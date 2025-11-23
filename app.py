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
    Processa o arquivo Excel no formato real - VERSÃO OTIMIZADA
    """
    print(f"Processando arquivo (visualização={para_visualizacao})...")
    
    try:
        # Ler arquivo em chunks para economizar memória
        excel_file = pd.ExcelFile(filepath)
        print(f"Planilhas encontradas: {excel_file.sheet_names}")
        
        # Processar planilha por planilha
        df_list = []
        total_linhas_lidas = 0
        
        for sheet_name in excel_file.sheet_names:
            print(f"Processando planilha '{sheet_name}'...")
            
            # Ler planilha
            df_temp = pd.read_excel(
                filepath, 
                sheet_name=sheet_name,
                usecols=None  # Ler todas as colunas primeiro
            )
            
            total_linhas_lidas += len(df_temp)
            print(f"  Linhas brutas: {len(df_temp)}")
            
            # Remover linhas completamente vazias IMEDIATAMENTE
            df_temp = df_temp.dropna(how='all')
            print(f"  Após remover vazias: {len(df_temp)}")
            
            if len(df_temp) > 0:
                df_list.append(df_temp)
            
            # Limpar memória
            del df_temp
        
        if not df_list:
            raise Exception("Nenhum dado válido encontrado nas planilhas")
        
        # Concatenar
        print("Concatenando planilhas...")
        df = pd.concat(df_list, ignore_index=True)
        del df_list  # Liberar memória
        
        print(f"Total após concatenar: {len(df)} linhas (de {total_linhas_lidas} originais)")
        
        # Identificar colunas BOX e STATUS
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
        
        print(f"Colunas identificadas: {colunas_map}")
        
        if 'BOX' not in colunas_map:
            raise Exception(f"Coluna BOX não encontrada. Colunas disponíveis: {list(df.columns)}")
        if 'STATUS' not in colunas_map:
            raise Exception(f"Coluna STATUS não encontrada. Colunas disponíveis: {list(df.columns)}")
        
        # Manter colunas essenciais + COD, SETOR, TIPO para análises
        colunas_manter = ['Box', 'Status']
        rename_map = {
            colunas_map['BOX']: 'Box',
            colunas_map['STATUS']: 'Status'
        }
        
        if 'COD' in colunas_map:
            colunas_manter.append('COD')
            rename_map[colunas_map['COD']] = 'COD'
        if 'SETOR' in colunas_map:
            colunas_manter.append('SETOR')
            rename_map[colunas_map['SETOR']] = 'SETOR'
        if 'TIPO' in colunas_map:
            colunas_manter.append('TIPO')
            rename_map[colunas_map['TIPO']] = 'TIPO'
        
        # Selecionar apenas colunas necessárias
        df = df[[col for col in df.columns if col in rename_map.keys()]]
        df = df.rename(columns=rename_map)
        
        print("Limpando dados...")
        
        # Remover linhas onde Box está vazio
        df = df[df['Box'].notna()]
        df = df[df['Box'] != '']
        
        # Converter Box para numérico
        df['Box'] = df['Box'].astype(str).str.strip()
        df['Box'] = df['Box'].str.replace(r'[^\d]', '', regex=True)
        df['Box'] = pd.to_numeric(df['Box'], errors='coerce')
        
        # Remover inválidos
        df = df.dropna(subset=['Box'])
        df['Box'] = df['Box'].astype(int)
        
        # Filtrar range válido
        df = df[(df['Box'] >= 1) & (df['Box'] <= 7000)]
        
        print(f"Documentos válidos: {len(df)}")
        
        if len(df) == 0:
            raise Exception("Nenhum documento com Box válido (1-7000) foi encontrado")
        
        # Limpar Status
        df['Status'] = df['Status'].fillna('SEM STATUS').astype(str).str.strip()
        
        # Limpar COD e SETOR se existirem
        if 'COD' in df.columns:
            df['COD'] = df['COD'].fillna('NÃO INFORMADO').astype(str).str.strip()
            df['COD'] = df['COD'].replace('', 'NÃO INFORMADO')
        
        if 'SETOR' in df.columns:
            df['SETOR'] = df['SETOR'].fillna('NÃO INFORMADO').astype(str).str.strip()
            df['SETOR'] = df['SETOR'].replace('', 'NÃO INFORMADO')
        
        if 'TIPO' in df.columns:
            df['TIPO'] = df['TIPO'].fillna('NÃO INFORMADO').astype(str).str.strip()
            df['TIPO'] = df['TIPO'].replace('', 'NÃO INFORMADO')
        
        # Estatísticas básicas
        status_unicos = df['Status'].unique()
        totais_por_status = df['Status'].value_counts().to_dict()
        total_geral = len(df)
        boxes_ocupados = df['Box'].nunique()
        
        print(f"✓ Total: {total_geral} docs, {boxes_ocupados} boxes, {len(status_unicos)} status")
        
        # Cores
        cores_status = gerar_cores_distintas(len(status_unicos))
        mapa_cores = dict(zip(status_unicos, cores_status))
        
        # Boxes data (apenas se necessário)
        boxes_data = {}
        dados_extras = {}
        
        if para_visualizacao:
            print("Gerando boxes_data...")
            grupos_box = df.groupby('Box')
            
            for box_num, docs_box in grupos_box:
                status_count = docs_box['Status'].value_counts().to_dict()
                total_box = len(docs_box)
                percentuais = {s: (c/total_box)*100 for s, c in status_count.items()}
                
                boxes_data[int(box_num)] = {
                    'total': total_box,
                    'situacoes': status_count,
                    'percentuais': percentuais
                }
            
            # Adicionar vazios
            for box_num in range(1, 7001):
                if box_num not in boxes_data:
                    boxes_data[box_num] = {'total': 0, 'situacoes': {}, 'percentuais': {}}
        else:
            # Relatórios: dados extras leves
            print("Gerando dados para relatórios...")
            status_por_box = df.groupby(['Box', 'Status']).size().reset_index(name='count')
            dados_extras['status_por_box'] = status_por_box.to_dict('records')
            
            # Análises adicionais: COD, SETOR, TIPO
            if 'COD' in df.columns:
                print("Gerando tabela de frequência para COD...")
                freq_cod = df['COD'].value_counts().reset_index()
                freq_cod.columns = ['Variável', 'Freq_Absoluta']
                freq_cod['Freq_Relativa'] = (freq_cod['Freq_Absoluta'] / len(df) * 100).round(2)
                freq_cod['Freq_Rel_Acumulada'] = freq_cod['Freq_Relativa'].cumsum().round(2)
                dados_extras['freq_cod'] = freq_cod.to_dict('records')
                print(f"  → {len(freq_cod)} códigos únicos encontrados")
                
                # COD mais frequente -> Análise por TIPO
                if len(freq_cod) > 0 and 'TIPO' in df.columns:
                    cod_mais_freq = freq_cod.iloc[0]['Variável']
                    df_cod_top = df[df['COD'] == cod_mais_freq]
                    print(f"  → COD mais frequente: '{cod_mais_freq}' ({len(df_cod_top)} documentos)")
                    
                    freq_tipo_cod = df_cod_top['TIPO'].value_counts().reset_index()
                    freq_tipo_cod.columns = ['Variável', 'Freq_Absoluta']
                    freq_tipo_cod['Freq_Relativa'] = (freq_tipo_cod['Freq_Absoluta'] / len(df_cod_top) * 100).round(2)
                    freq_tipo_cod['Freq_Rel_Acumulada'] = freq_tipo_cod['Freq_Relativa'].cumsum().round(2)
                    
                    dados_extras['freq_tipo_por_cod_top'] = {
                        'cod': str(cod_mais_freq),
                        'total_docs': int(freq_cod.iloc[0]['Freq_Absoluta']),
                        'dados': freq_tipo_cod.to_dict('records')
                    }
                    print(f"  → {len(freq_tipo_cod)} tipos únicos no COD top")
            else:
                print("  ⚠ Coluna COD não encontrada, análise ignorada")
            
            if 'SETOR' in df.columns:
                print("Gerando tabela de frequência para SETOR...")
                freq_setor = df['SETOR'].value_counts().reset_index()
                freq_setor.columns = ['Variável', 'Freq_Absoluta']
                freq_setor['Freq_Relativa'] = (freq_setor['Freq_Absoluta'] / len(df) * 100).round(2)
                freq_setor['Freq_Rel_Acumulada'] = freq_setor['Freq_Relativa'].cumsum().round(2)
                dados_extras['freq_setor'] = freq_setor.to_dict('records')
                print(f"  → {len(freq_setor)} setores únicos encontrados")
                
                # SETOR mais frequente -> Análise por TIPO
                if len(freq_setor) > 0 and 'TIPO' in df.columns:
                    setor_mais_freq = freq_setor.iloc[0]['Variável']
                    df_setor_top = df[df['SETOR'] == setor_mais_freq]
                    print(f"  → SETOR mais frequente: '{setor_mais_freq}' ({len(df_setor_top)} documentos)")
                    
                    freq_tipo_setor = df_setor_top['TIPO'].value_counts().reset_index()
                    freq_tipo_setor.columns = ['Variável', 'Freq_Absoluta']
                    freq_tipo_setor['Freq_Relativa'] = (freq_tipo_setor['Freq_Absoluta'] / len(df_setor_top) * 100).round(2)
                    freq_tipo_setor['Freq_Rel_Acumulada'] = freq_tipo_setor['Freq_Relativa'].cumsum().round(2)
                    
                    dados_extras['freq_tipo_por_setor_top'] = {
                        'setor': str(setor_mais_freq),
                        'total_docs': int(freq_setor.iloc[0]['Freq_Absoluta']),
                        'dados': freq_tipo_setor.to_dict('records')
                    }
                    print(f"  → {len(freq_tipo_setor)} tipos únicos no SETOR top")
            else:
                print("  ⚠ Coluna SETOR não encontrada, análise ignorada")
        
        print("✓ Processamento concluído!")
        
        resultado = {
            'boxes_data': boxes_data,
            'colunas_situacoes': list(status_unicos),
            'mapa_cores': mapa_cores,
            'totais_por_situacao': totais_por_status,
            'total_geral': total_geral,
            'boxes_ocupados': boxes_ocupados,
            'total_boxes': boxes_ocupados
        }
        
        if not para_visualizacao:
            resultado.update(dados_extras)
        
        return resultado
        
    except Exception as e:
        print(f"✗ ERRO: {str(e)}")
        print(traceback.format_exc())
        raise

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    print("=== ROTA /upload CHAMADA ===")
    
    from flask import make_response
    
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
        
        print("Iniciando processamento para visualização...")
        # Processar para visualização (completo)
        dados = processar_excel_novo_formato(filepath, para_visualizacao=True)
        print(f"Dados processados: {dados['total_geral']} documentos")
        
        # Remover arquivo temporário
        try:
            os.remove(filepath)
        except:
            pass
        
        # Criar resposta com headers explícitos
        response = make_response(jsonify(dados))
        response.headers['Content-Type'] = 'application/json; charset=utf-8'
        return response
        
    except Exception as e:
        print(f"ERRO: {str(e)}")
        print(traceback.format_exc())
        error_response = make_response(jsonify({'error': str(e)}))
        error_response.headers['Content-Type'] = 'application/json; charset=utf-8'
        return error_response, 500

@app.route('/visualizar')
def visualizar():
    return render_template('visualizacao.html')

@app.route('/upload-relatorios', methods=['POST'])
def upload_relatorios():
    print("=== ROTA /upload-relatorios CHAMADA ===")
    
    # Adicionar headers CORS para debug
    from flask import make_response
    
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
        
        print("Iniciando processamento...")
        # Processar para relatórios (sem boxes_data completo)
        dados = processar_excel_novo_formato(filepath, para_visualizacao=False)
        print(f"Dados processados com sucesso: {len(dados)} chaves")
        
        # Remover arquivo temporário
        try:
            os.remove(filepath)
            print("Arquivo temporário removido")
        except Exception as e:
            print(f"Aviso: não foi possível remover arquivo temporário: {e}")
        
        # Criar resposta com headers explícitos
        response = make_response(jsonify(dados))
        response.headers['Content-Type'] = 'application/json; charset=utf-8'
        print("Retornando resposta JSON")
        return response
        
    except Exception as e:
        print(f"ERRO CRÍTICO: {str(e)}")
        print(traceback.format_exc())
        error_response = make_response(jsonify({'error': str(e)}))
        error_response.headers['Content-Type'] = 'application/json; charset=utf-8'
        return error_response, 500

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
