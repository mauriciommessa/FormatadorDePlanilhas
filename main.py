import pandas as pd

#Função para formatar a data
def formatar_data(data):
    return data.strftime('%d/%m/%Y')


def formatar_valor(valor):
    return f'R$ {valor:.2f}'


tabela = pd.read_excel("exemploAvarias.xlsx")

tabela['dt_servico'] = tabela['dt_servico'].apply(formatar_data)

tabelaAgrupado = tabela.groupby(['placa_ou_id_veiculo', 
                                  'servico 1-Manutencao Corretiva 2-Manutecao Preventiva 3-Sinistro 4-Ordem Servico 5-AVARIAS RECUPERACAO/DEVOLUCAO 6-TROCA PNEUS 7-TAXA ADMINISTRATIVA', 'dt_servico', 'numero_ordem_servico']).agg({'descricao': ':'.join, 'valor': 'sum'}).reset_index()

tabelaAgrupado['valor_total'] = tabelaAgrupado.groupby(['placa_ou_id_veiculo', 'servico 1-Manutencao Corretiva 2-Manutecao Preventiva 3-Sinistro 4-Ordem Servico 5-AVARIAS RECUPERACAO/DEVOLUCAO 6-TROCA PNEUS 7-TAXA ADMINISTRATIVA'])['valor'].transform('sum')

tabelaAgrupado['valor_total'] = tabelaAgrupado['valor_total'].apply(formatar_valor)

print(tabelaAgrupado)

#Salvar os resultados em um novo arquivo Excel
tabelaAgrupado.to_excel('novo_arquivo1.xlsx', index=False)