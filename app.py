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

def processar_excel_novo_formato(filepath, para_visualizacao=True):
    """
    Processa o arquivo Excel no formato real:
    SETOR | DATA | STATUS | BOX | COD | TIPO | ELIM_PREVISTA
    """
    print(f"Processando arquivo (visualização={para_visualizacao})...")
    
    try:
        # Ler todas as planilhas
        excel_file = pd.ExcelFile(filepath)
        print(f"Planilhas encontradas: {excel_file.sheet_names}")
        
        # Concatenar todas as planilhas
        dfs = []
        for sheet_name in excel_file.sheet_names:
            df_temp = pd.read_excel(filepath, sheet_name=sheet_name)
            print(f"Planilha '{sheet_name}': {len(df_temp)} linhas")
            dfs.append(df_temp)
        
        df = pd.concat(dfs, ignore_index=True)
        print(f"Total de linhas após concatenar: {len(df)}")
        print(f"Colunas encontradas: {list(df.columns)}")
        
        # Identificar colunas (case-insensitive)
        colunas_map = {}
        for col in df.columns:
            col_upper = str(col).upper().strip()
            if 'BOX' in col_upper:
                colunas_map['BOX'] = col
            elif 'STATUS' in col_upper or 'SITUAÇÃO' in col_upper or 'SITUACAO' in col_upper:
                colunas_map['STATUS'] = col
            elif 'SETOR' in col_upper:
                colunas_map['SETOR'] = col
            elif 'COD' in col_upper and 'TIPO' not in col_upper:
                colunas_map['COD'] = col
            elif 'TIPO' in col_upper:
                colunas_map['TIPO'] = col
        
        print(f"Mapeamento de colunas: {colunas_map}")
        
        # Verificar se encontrou as colunas essenciais
        if 'BOX' not in colunas_map or 'STATUS' not in colunas_map:
            raise Exception(f"Colunas essenciais não encontradas. Encontradas: {list(df.columns)}")
        
        # Renomear colunas para padrão
        df = df.rename(columns={
            colunas_map.get('BOX'): 'Box',
            colunas_map.get('STATUS'): 'Status'
        })
        
        # Converter coluna Box para numérico (remover textos, espaços, etc)
        df['Box'] = df['Box'].astype(str).str.replace(r'[^\d]', '', regex=True)
        df['Box'] = pd.to_numeric(df['Box'], errors='coerce')
        
        # Remover linhas sem Box válido
        df = df.dropna(subset=['Box'])
        df['Box'] = df['Box'].astype(int)
        
        # Filtrar boxes de 1 a 7000
        df = df[(df['Box'] >= 1) & (df['Box'] <= 7000)]
        print(f"Documentos em boxes válidos (1-7000): {len(df)}")
        
        # Preencher Status vazios
        df['Status'] = df['Status'].fillna('SEM STATUS')
        
        # Obter lista de status únicos
        status_unicos = df['Status'].unique()
        print(f"Status encontrados: {list(status_unicos)}")
        
        # Contar documentos por status
        totais_por_status = df['Status'].value_counts().to_dict()
        
        # Total de documentos
        total_geral = len(df)
        
        # Contar boxes únicos ocupados
        boxes_ocupados = df['Box'].nunique()
        total_boxes = len(df['Box'].unique())
        
        print(f"Total de documentos: {total_geral}")
        print(f"Boxes ocupados: {boxes_ocupados}")
        print(f"Status: {totais_por_status}")
        
        # Gerar cores
        cores_status = gerar_cores_distintas(len(status_unicos))
        mapa_cores = dict(zip(status_unicos, cores_status))
        
        # Se for para visualização, criar boxes_data completo
        boxes_data = {}
        if para_visualizacao:
            print("Gerando estrutura de visualização de boxes...")
            
            # Agrupar documentos por box
            grupos_box = df.groupby('Box')
            
            # Processar cada box
            for box_num in range(1, 7001):
                if box_num in grupos_box.groups:
                    docs_box = grupos_box.get_group(box_num)
                    
                    # Contar por status
                    status_count = docs_box['Status'].value_counts().to_dict()
                    total_box = len(docs_box)
                    
                    # Calcular percentuais
                    percentuais = {}
                    for status, count in status_count.items():
                        percentuais[status] = (count / total_box) * 100
                    
                    boxes_data[box_num] = {
                        'total': total_box,
                        'situacoes': status_count,
                        'percentuais': percentuais
                    }
                else:
                    boxes_data[box_num] = {
                        'total': 0,
                        'situacoes': {},
                        'percentuais': {}
                    }
        
        # Para relatórios, incluir dados adicionais
        dados_extras = {}
        if not para_visualizacao:
            # Contar documentos por box (para estatísticas detalhadas)
            docs_por_box = df.groupby('Box').size().to_dict()
            dados_extras['docs_por_box'] = docs_por_box
            
            # Estatísticas por status e box
            status_por_box = df.groupby(['Box', 'Status']).size().reset_index(name='count')
            dados_extras['status_por_box'] = status_por_box.to_dict('records')
        
        print("Processamento concluído!")
        
        resultado = {
            'boxes_data': boxes_data,
            'colunas_situacoes': list(status_unicos),
            'mapa_cores': mapa_cores,
            'totais_por_situacao': totais_por_status,
            'total_geral': total_geral,
            'boxes_ocupados': boxes_ocupados,
            'total_boxes': total_boxes
        }
        
        # Adicionar dados extras para relatórios
        if not para_visualizacao:
            resultado.update(dados_extras)
        
        return resultado
        
    except Exception as e:
        print(f"Erro no processamento: {str(e)}")
        print(traceback.format_exc())
        raise

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    print("=== ROTA /upload CHAMADA ===")
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['file']
        print(f"Arquivo recebido: {file.filename}")
        
        if file.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Tipo de arquivo não permitido. Use .xls ou .xlsx'}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Processar para visualização (completo)
        dados = processar_excel_novo_formato(filepath, para_visualizacao=True)
        
        # Remover arquivo temporário
        try:
            os.remove(filepath)
        except:
            pass
        
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
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['file']
        print(f"Arquivo recebido: {file.filename}")
        
        if file.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Tipo de arquivo não permitido. Use .xls ou .xlsx'}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Processar para relatórios (sem boxes_data completo)
        dados = processar_excel_novo_formato(filepath, para_visualizacao=False)
        
        # Remover arquivo temporário
        try:
            os.remove(filepath)
        except:
            pass
        
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
    return jsonify({'status': 'ok'}), 200

if __name__ == '__main__':
    print("=== INICIANDO SERVIDOR ===")
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
