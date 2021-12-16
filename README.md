# calculadora_frete
Script feito para limpar relatórios de fatura do sistema da empresa, fornecendo informação do que foi cobrado dos clientes na entrega dos produtos e em seguida recalcular o que deveria ter sido cobrado e indicar os motivos das divergências

Toma como input os seguintes arquivos:

-Relatório descritivo: Relatório que contém dados das remessas como valor da mercadoria, cep de origem, cep destino, valor total cobrado, id do pedido.

-Relatório demonstrativo: Relatório que contém dados das remessas como valor de cada componente cobrado (Advalorem, frete principal, icms, outros) e a modalidade do produto

-Tabela comercial: Tabela de preços do cliente

-Cidaten: faixas de CEP atendidas pela empresa e quais taxas são cobradas por faixa.

Além disso são imputadas as generalidades, que são percentuais de taxas a serem cobradas e variam de cliente para cliente.

------------------------------------------------------------------------------------------------

Após ler os arquivos o script pega as informações relevantes para calculo do valor de serviço de cada remessa e recalcula o que deveria ter sido cobrado através dos parâmetros informados
