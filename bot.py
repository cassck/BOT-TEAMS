from flask import Flask, request, jsonify, render_template
import pandas as pd
import os
import openpyxl
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configurações
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

# Caminho para o arquivo pai (agora lido de uma variável de ambiente)
CAMINHO_PAI = os.environ.get('CAMINHO_PAI_SAP', r"C:\Users\erick.pereira\Downloads\Solicitação de Cadastros de Material - SAP.xlsx")

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/processar', methods=['POST'])
def processar_arquivo():
    if 'arquivo_filho' not in request.files:
        return jsonify({'success': False, 'message': 'Nenhum arquivo enviado'})

    file = request.files['arquivo_filho']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'Nome de arquivo inválido'})

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        caminho_filho = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(caminho_filho)

        try:
            success = adicionar_linhas_excel_existente(CAMINHO_PAI, caminho_filho)
            if success:
                return jsonify({
                    'success': True,
                    'message': 'Linhas adicionadas com sucesso ao arquivo pai!'
                })
            else:
                return jsonify({
                    'success': False,
                    'message': 'Falha ao processar o arquivo'
                })
        except Exception as e:
            return jsonify({
                'success': False,
                'message': f'Erro durante o processamento: {str(e)}'
            })
        finally:
            if os.path.exists(caminho_filho):
                os.remove(caminho_filho)

    return jsonify({'success': False, 'message': 'Tipo de arquivo não permitido'})

def ler_arquivo_excel_sem_cabecalho(caminho_arquivo, skiprows=2, sheet_name=0):
    try:
        if not os.path.exists(caminho_arquivo):
            raise FileNotFoundError(f"O arquivo '{caminho_arquivo}' não foi encontrado.")

        df = pd.read_excel(caminho_arquivo, skiprows=skiprows, sheet_name=sheet_name)
        return df

    except FileNotFoundError as e:
        print(f"Erro: {e}")
        return None
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return None


def adicionar_linhas_excel_existente(caminho_arquivo_pai, caminho_arquivo_filho, skiprows_filho=2, sheet_name_pai=0, sheet_name_filho=0):
    try:
        # 1. Ler o arquivo filho (com tratamento de cabeçalho)
        df_filho = ler_arquivo_excel_sem_cabecalho(caminho_arquivo_filho, skiprows=skiprows_filho, sheet_name=sheet_name_filho)
        if df_filho is None:
            return False

        df_filho.columns = df_filho.iloc[0]
        df_filho = df_filho.iloc[1:].reset_index(drop=True)

        # 2. Carregar o arquivo pai *usando openpyxl* (para modificação)
        workbook = openpyxl.load_workbook(caminho_arquivo_pai)

        # 3. Selecionar a planilha correta (usando o nome ou índice)
        if isinstance(sheet_name_pai, int):
            sheet = workbook.worksheets[sheet_name_pai]  # Acessa pela posição (índice)
        else:
            sheet = workbook[sheet_name_pai]            # Acessa pelo nome

        #Ler o arquivo pai usando o pandas, para verificar os cabeçalhos
        df_pai = pd.read_excel(caminho_arquivo_pai, sheet_name=sheet_name_pai)


        #Verificar se o df_pai e df_filho tem o mesmo numero de colunas
        if df_pai.shape[1] != df_filho.shape[1]:
            raise ValueError("Os arquivos Excel não têm o mesmo número de colunas. Não é possível concatenar.")

        #Verificar se o cabeçalho é igual, se não for, ajusta o do arquivo filho
        if list(df_pai.columns) != list(df_filho.columns):
            print("Aviso: Os cabeçalhos dos arquivos são diferentes. Ajustando o cabeçalho do arquivo filho...")
            #Tentar encontrar um mapeamento de colunas
            colunas_em_comum = [col for col in df_pai.columns if col in df_filho.columns]
            if len(colunas_em_comum) > 0:
                df_filho = df_filho[colunas_em_comum]  # Mantém apenas colunas em comum
                df_filho.columns = df_pai.columns[:len(colunas_em_comum)] #Alinha o cabeçalho do filho com o do pai (primeiras colunas)
            else:
                #Se não houver colunas em comum, usar a estratégia anterior
                print("Aviso: Não há colunas em comum. Alinhando colunas por posição")
                if df_pai.shape[1] >= df_filho.shape[1]:
                    df_filho.columns = df_pai.columns[:df_filho.shape[1]] #Alinha o cabeçalho do filho com as primeiras colunas do pai.
                else:
                    #Adiciona colunas vazias ao df_filho se o df_pai tiver menos colunas
                    for i in range (df_filho.shape[1], df_pai.shape[1]):
                        df_filho[f'Coluna_{i}'] = None  # Adiciona colunas vazias
                    df_filho.columns = df_pai.columns


        # 4. Adicionar os dados do DataFrame filho à planilha
        for row in df_filho.values.tolist():  # Itera pelas linhas do DataFrame filho
            sheet.append(row)  # Adiciona cada linha à planilha

        # 5. Salvar as alterações *no mesmo arquivo pai*
        workbook.save(caminho_arquivo_pai)
        return True

    except FileNotFoundError:
        print(f"Erro: Um dos arquivos não foi encontrado.")
        return False
    except ValueError as ve:
        print(f"Erro de validação: {ve}")
        return False
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
        return False

if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True)