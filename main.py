import pandas as pd

tabela = pd.read_excel("exemploAvarias.xlsx")

tabelaAgrupado = tabela.groupby(['placa_ou_id_veiculo',
                                 'servico 1-Manutencao Corretiva 2-Manutecao Preventiva 3-Sinistro 4-Ordem Servico 5-AVARIAS RECUPERACAO/DEVOLUCAO 6-TROCA PNEUS 7-TAXA ADMINISTRATIVA',
                                 'dt_servico',
                                 'numero_ordem_servico']).agg({'valor': 'sum', 'descricao': ' / '.join}).reset_index()

print(tabelaAgrupado)

tabelaAgrupado.to_excel('novo_arquivo.xlsx', index=False)
