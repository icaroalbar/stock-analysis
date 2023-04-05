# Documentação do código

Este código tem como objetivo gerar relatórios de desempenho financeiro para empresas. Os dados são obtidos da biblioteca yfinance e salvos em um arquivo Excel.

## Bibliotecas utilizadas

- openpyxl: utilizada para salvar os relatórios em um arquivo Excel.
- empresas: lista de códigos de ações das empresas que serão avaliadas.
- yfinance: utilizada para obter dados financeiros das empresas.
- datetime: utilizada para obter a data atual.
- calendar: utilizada para obter o último dia do mês.
- pandas: utilizada para manipular os dados obtidos e criar o relatório.
- playsound: utilizada para reproduzir um som ao final da execução (descomentada).

## Funções

### obter_ultimo_dia_do_mes(ano, mes)

Esta função retorna o último dia do mês informado.

### obter_data_fim_disponivel_para_acao(acao, ano, mes)

Esta função retorna a última data disponível para a ação no mês informado.

### obter_dados_do_periodo(acao, ano, mes)

Esta função retorna os dados financeiros da ação no período informado.

### calcular_retorno_e_variacao_percentual(dados)

Esta função calcula o retorno e a variação percentual dos dados financeiros da ação.

### gerar_relatorio(acao, ano, mes)

Esta função gera um relatório para a ação e mês informados. O relatório inclui informações como o retorno e a variação percentual da ação no período.

### gerar_relatorio_para_meses_anteriores(acao, n_anos)

Esta função gera um relatório para a ação nos últimos n_anos informados. A função retorna um DataFrame com os dados do relatório.

### salvar_relatorios_em_excel()

Esta função salva os relatórios gerados para cada empresa em um arquivo Excel. O nome do arquivo inclui a data atual. A função também verifica se há arquivos antigos com o mesmo nome e os exclui. Se não houver relatórios disponíveis, a função ainda cria o arquivo Excel, mas com uma mensagem indicando que não há relatórios disponíveis. A planilha gerada inclui uma guia para as empresas que não possuem relatório e as células da primeira linha estão congeladas para facilitar a visualização dos dados. Além disso, a data e hora de geração da planilha são adicionadas na primeira coluna.

## Execução

Por fim, a função `salvar_relatorios_em_excel()` é chamada para gerar e salvar os relatórios em um arquivo Excel. Opcionalmente, ao final da execução, um som pode ser reproduzido para indicar o sucesso da geração dos relatórios (comentado no código).
