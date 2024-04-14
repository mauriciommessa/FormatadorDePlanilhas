import os
import pandas as pd

def formatar_data(data):
    if isinstance(data, pd.Timestamp):  # Verificar se é um objeto datetime
        return data.strftime('%d/%m/%Y')
    else:
        return ''

def formatar_valor(valor):
    if pd.isnull(valor):  # Verificar se o valor é NaN
        return ''  # Se for NaN, retornar uma string vazia
    else:
        return f'{valor:.2f} '

def processar_planilha(file_path):
    try:
        # Carregar os dados do arquivo Excel
        tabela = pd.read_excel(file_path)

        # Converter a coluna 'dt_servico' para o tipo datetime
        tabela['dt_servico'] = pd.to_datetime(tabela['dt_servico'], errors='coerce')

        # Converter a coluna 'valor' para float
        tabela['valor'] = pd.to_numeric(tabela['valor'], errors='coerce')

        # Formatar a data
        tabela['dt_servico'] = tabela['dt_servico'].apply(formatar_data)

        # Formatar o valor em reais (R$) e armazenar em uma nova coluna
        tabela['valor_formatado'] = tabela['valor'].apply(formatar_valor)

        # Agrupar os dados da tabela ignorando a data e mantendo a primeira data de cada grupo
        tabelaAgrupado = tabela.groupby(['placa_ou_id_veiculo',
                                         'servico 1-Manutencao Corretiva 2-Manutecao Preventiva 3-Sinistro 4-Ordem Servico 5-AVARIAS RECUPERACAO/DEVOLUCAO 6-TROCA PNEUS 7-TAXA ADMINISTRATIVA',
                                         'numero_ordem_servico']).agg({'valor': 'sum', 'descricao': ' / '.join, 'dt_servico': 'first'}).reset_index()

        # Salvar os resultados em um novo arquivo Excel
        new_file_path = os.path.splitext(file_path)[0] + '_formatado.xlsx'
        tabelaAgrupado.to_excel(new_file_path, index=False)
        print(f"Planilha {file_path} processada e salva em {new_file_path}")
    except Exception as e:
        print(f"Erro ao processar a planilha {file_path}: {e}")


folder_path = './'


for file_name in os.listdir(folder_path):
    # Construir o caminho completo do arquivo
    file_path = os.path.join(folder_path, file_name)


    if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        # Processar a planilha
        processar_planilha(file_path)
