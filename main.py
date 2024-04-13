import pandas as pd

arquivoEntrada = "planilha.xlsx"
arquivoSaida = "arquivoOrganizado.xlsx"


def formatar_data(data):
    # formata a data para dd/mm/yyyy
    if isinstance(data, pd.Timestamp):  # Verificar se é um objeto datetime
        return data.strftime('%d/%m/%Y')
    else:
        return ''


def formatar_valor(valor):
    # formata o valor para R$xx.xx
    return f'R$ {valor:.2f}'


def carregar_dados(arquivo):
    # carrega os dados do arquivo Excel
    try:
        tabela = pd.read_excel(arquivo)
    except FileNotFoundError:
        print(f'O arquivo {arquivo} não foi encontrado.')
        return None

    return tabela


def processar_dados(tabela):
    # processa os dados da tabela
    tabela['dt_servico'] = pd.to_datetime(tabela['dt_servico'], errors='coerce')
    tabela = tabela.dropna(subset=['dt_servico'])
    tabela['valor'] = pd.to_numeric(tabela['valor'], errors='coerce')
    tabela['dt_servico'] = tabela['dt_servico'].apply(formatar_data)
    tabela['valor'] = tabela['valor'].apply(formatar_valor)

    return tabela


def agrupar_dados(tabela):
    # agrupa os dados da tabela
    tabelaAgrupado = tabela.groupby(['placa_ou_id_veiculo',
                                     'servico 1-Manutencao Corretiva 2-Manutecao Preventiva 3-Sinistro 4-Ordem Servico 5-AVARIAS RECUPERACAO/DEVOLUCAO 6-TROCA PNEUS 7-TAXA ADMINISTRATIVA',
                                     'numero_ordem_servico',
                                     'dt_servico']).agg({'valor': 'sum', 'descricao': ' / '.join}).reset_index()

    return tabelaAgrupado


def salvar_dados(tabela, arquivo):
    # salva os dados processados em um novo arquivo Excel
    tabela.to_excel(arquivo, index=False)


def main():
    tabela = carregar_dados(arquivoEntrada)
    if tabela is not None:
        tabela = processar_dados(tabela)
        tabelaAgrupado = agrupar_dados(tabela)
        print(tabelaAgrupado)
        salvar_dados(tabelaAgrupado, arquivoSaida)


if __name__ == "__main__":
    main()
