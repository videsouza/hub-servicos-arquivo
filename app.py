from flask import Flask, render_template, request, jsonify
import pandas as pd
import colorsys
import os
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def gerar_cores_distintas(n):
    """Gera n cores distintas no espectro HSV"""
    cores = []
    for i in range(n):
        hue = i / n
        rgb = colorsys.hsv_to_rgb(hue, 0.7, 0.9)
        hex_color = '#{:02x}{:02x}{:02x}'.format(
            int(rgb[0] * 255),
            int(rgb[1] * 255),
            int(rgb[2] * 255)
        )
        cores.append(hex_color)
    return cores

def processar_excel(filepath):
    """Processa o arquivo Excel e retorna os dados estruturados"""
    try:
        # Ler arquivo Excel - cabeçalhos na linha 3 (índice 2)
        df = pd.read_excel(filepath, header=2)
        
        # Renomear primeira coluna para 'Box' se necessário
        if df.columns[0] != 'Box':
            df = df.rename(columns={df.columns[0]: 'Box'})
        
        # Converter coluna Box para numérico
        df['Box'] = pd.to_numeric(df['Box'], errors='coerce')
        
        # Remover linhas onde Box não é número válido
        df = df.dropna(subset=['Box'])
        df['Box'] = df['Box'].astype(int)
        
        # Filtrar apenas boxes de 1 a 7000
        df = df[(df['Box'] >= 1) & (df['Box'] <= 7000)]
        
        # Identificar colunas de situações (excluir 'Box' e 'Total')
        colunas_situacoes = [col for col in df.columns if col not in ['Box', 'Total', 'total', 'TOTAL']]
        
        # Preencher NaN com 0 nas colunas de situações
        for col in colunas_situacoes:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        # Criar estrutura de dados para visualização
        boxes_data = {}
        
        for box_num in range(1, 7001):
            box_row = df[df['Box'] == box_num]
            
            if len(box_row) == 0:
                boxes_data[box_num] = {
                    'total': 0,
                    'situacoes': {},
                    'percentuais': {}
                }
            else:
                box_row = box_row.iloc[0]
                situacoes = {}
                percentuais = {}
                total_box = 0
                
                for col in colunas_situacoes:
                    count = int(box_row[col])
                    if count > 0:
                        situacoes[col] = count
                        total_box += count
                
                if total_box > 0:
                    for sit, count in situacoes.items():
                        percentuais[sit] = (count / total_box) * 100
                
                boxes_data[box_num] = {
                    'total': total_box,
                    'situacoes': situacoes,
                    'percentuais': percentuais
                }
        
        # Calcular estatísticas
        totais_por_situacao = {}
        for col in colunas_situacoes:
            total = df[col].sum()
            totais_por_situacao[col] = int(total)
        
        total_geral = sum(totais_por_situacao.values())
        boxes_ocupados = len(df[df[colunas_situacoes].sum(axis=1) > 0])
        
        # Gerar cores
        cores_situacoes = gerar_cores_distintas(len(colunas_situacoes))
        mapa_cores = dict(zip(colunas_situacoes, cores_situacoes))
        
        return {
            'boxes_data': boxes_data,
            'colunas_situacoes': colunas_situacoes,
            'mapa_cores': mapa_cores,
            'totais_por_situacao': totais_por_situacao,
            'total_geral': total_geral,
            'boxes_ocupados': boxes_ocupados,
            'total_boxes': len(df)
        }
    except Exception as e:
        raise Exception(f'Erro ao processar Excel: {str(e)}')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Tipo de arquivo não permitido'}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Processar arquivo
        dados = processar_excel(filepath)
        
        # Remover arquivo temporário
        try:
            os.remove(filepath)
        except:
            pass
        
        return jsonify(dados)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/visualizar')
def visualizar():
    return render_template('visualizacao.html')

@app.route('/upload-relatorios', methods=['POST'])
def upload_relatorios():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Tipo de arquivo não permitido'}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Processar arquivo
        dados = processar_excel(filepath)
        
        # Remover arquivo temporário
        try:
            os.remove(filepath)
        except:
            pass
        
        return jsonify(dados)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/relatorios')
def relatorios():
    return render_template('relatorios.html')

# Rota de health check para o Render
@app.route('/health')
def health():
    return jsonify({'status': 'ok'}), 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
