from flask import Flask, render_template, request, jsonify
import pandas as pd
import colorsys
import os
from werkzeug.utils import secure_filename
import tempfile
import traceback

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

def processar_excel_visualizacao(filepath):
    """Processa o arquivo Excel para visualização de boxes (completo)"""
    print("Processando para visualização...")
    
    # Ler arquivo Excel
    df = pd.read_excel(filepath, header=2)
    
    # Renomear primeira coluna para 'Box'
    if df.columns[0] != 'Box':
        df = df.rename(columns={df.columns[0]: 'Box'})
    
    # Converter e limpar
    df['Box'] = pd.to_numeric(df['Box'], errors='coerce')
    df = df.dropna(subset=['Box'])
    df['Box'] = df['Box'].astype(int)
    df = df[(df['Box'] >= 1) & (df['Box'] <= 7000)]
    
    # Identificar colunas de situações
    colunas_situacoes = [col for col in df.columns if col not in ['Box', 'Total', 'total', 'TOTAL']]
    
    # Preencher NaN
    for col in colunas_situacoes:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    
    print(f"Processando {len(df)} boxes...")
    
    # Criar estrutura de dados para visualização
    boxes_data = {}
    
    # Processar apenas boxes que existem no arquivo
    for _, row in df.iterrows():
        box_num = int(row['Box'])
        situacoes = {}
        percentuais = {}
        total_box = 0
        
        for col in colunas_situacoes:
            count = int(row[col])
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
    
    # Adicionar boxes vazios (1-7000) de forma eficiente
    print("Adicionando boxes vazios...")
    for box_num in range(1, 7001):
        if box_num not in boxes_data:
            boxes_data[box_num] = {
                'total': 0,
                'situacoes': {},
                'percentuais': {}
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
    
    print("Processamento concluído!")
    
    return {
        'boxes_data': boxes_data,
        'colunas_situacoes': colunas_situacoes,
        'mapa_cores': mapa_cores,
        'totais_por_situacao': totais_por_situacao,
        'total_geral': total_geral,
        'boxes_ocupados': boxes_ocupados,
        'total_boxes': len(df)
    }

def processar_excel_relatorios(filepath):
    """Processa o arquivo Excel apenas para relatórios (sem boxes_data completo)"""
    print("Processando para relatórios (versão leve)...")
    
    # Ler arquivo Excel
    df = pd.read_excel(filepath, header=2)
    
    # Renomear primeira coluna para 'Box'
    if df.columns[0] != 'Box':
        df = df.rename(columns={df.columns[0]: 'Box'})
    
    # Converter e limpar
    df['Box'] = pd.to_numeric(df['Box'], errors='coerce')
    df = df.dropna(subset=['Box'])
    df['Box'] = df['Box'].astype(int)
    df = df[(df['Box'] >= 1) & (df['Box'] <= 7000)]
    
    # Identificar colunas de situações
    colunas_situacoes = [col for col in df.columns if col not in ['Box', 'Total', 'total', 'TOTAL']]
    
    # Preencher NaN
    for col in colunas_situacoes:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    
    # Calcular estatísticas (não precisa do boxes_data completo)
    totais_por_situacao = {}
    for col in colunas_situacoes:
        total = df[col].sum()
        totais_por_situacao[col] = int(total)
    
    total_geral = sum(totais_por_situacao.values())
    boxes_ocupados = len(df[df[colunas_situacoes].sum(axis=1) > 0])
    
    # Gerar cores
    cores_situacoes = gerar_cores_distintas(len(colunas_situacoes))
    mapa_cores = dict(zip(colunas_situacoes, cores_situacoes))
    
    # Criar boxes_data vazio (relatórios não precisam disso)
    print("Processamento de relatórios concluído!")
    
    return {
        'boxes_data': {},  # Vazio para economizar memória
        'colunas_situacoes': colunas_situacoes,
        'mapa_cores': mapa_cores,
        'totais_por_situacao': totais_por_situacao,
        'total_geral': total_geral,
        'boxes_ocupados': boxes_ocupados,
        'total_boxes': len(df)
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    print("=== ROTA /upload CHAMADA ===")
    try:
        if 'file' not in request.files:
            print("Erro: Nenhum arquivo no request")
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['file']
        print(f"Arquivo recebido: {file.filename}")
        
        if file.filename == '':
            print("Erro: Nome do arquivo vazio")
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        if not allowed_file(file.filename):
            print("Erro: Tipo de arquivo não permitido")
            return jsonify({'error': 'Tipo de arquivo não permitido. Use .xls ou .xlsx'}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        print(f"Salvando em: {filepath}")
        file.save(filepath)
        
        print("Processando arquivo para visualização...")
        dados = processar_excel_visualizacao(filepath)
        print(f"Processamento concluído. Total de documentos: {dados['total_geral']}")
        
        # Remover arquivo temporário
        try:
            os.remove(filepath)
            print("Arquivo temporário removido")
        except Exception as e:
            print(f"Erro ao remover arquivo: {e}")
        
        return jsonify(dados)
        
    except Exception as e:
        print(f"ERRO: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

@app.route('/visualizar')
def visualizar():
    return render_template('visualizacao.html')

@app.route('/upload-relatorios', methods=['POST'])
def upload_relatorios():
    print("=== ROTA /upload-relatorios CHAMADA ===")
    try:
        if 'file' not in request.files:
            print("Erro: Nenhum arquivo no request")
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['file']
        print(f"Arquivo recebido: {file.filename}")
        
        if file.filename == '':
            print("Erro: Nome do arquivo vazio")
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        if not allowed_file(file.filename):
            print("Erro: Tipo de arquivo não permitido")
            return jsonify({'error': 'Tipo de arquivo não permitido. Use .xls ou .xlsx'}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        print(f"Salvando em: {filepath}")
        file.save(filepath)
        
        print("Processando arquivo para relatórios...")
        dados = processar_excel_relatorios(filepath)
        print(f"Processamento concluído. Total de documentos: {dados['total_geral']}")
        
        # Remover arquivo temporário
        try:
            os.remove(filepath)
            print("Arquivo temporário removido")
        except Exception as e:
            print(f"Erro ao remover arquivo: {e}")
        
        return jsonify(dados)
        
    except Exception as e:
        print(f"ERRO: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

@app.route('/relatorios')
def relatorios():
    return render_template('relatorios.html')

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'message': 'Servidor funcionando'}), 200

if __name__ == '__main__':
    print("=== INICIANDO SERVIDOR ===")
    print(f"Pasta de upload: {app.config['UPLOAD_FOLDER']}")
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
