import pandas as pd
import math
pd.options.mode.chained_assignment = None
import time
import numpy as np

time_i = time.time()
'''
Script para calcular e identificar origem de divergências em faturas.
Ele toma como input parâmetros de generalidades, tabela frete/peso, cidaten. Futuramente: iata, icms

Desenvolvedor: Daniel Watanabe
'''
#-----------------FUNCOES------------------#

#Função para realizar todos os cálculos de componentes
def calcular_componentes(cep_origem,cep_destino,peso,valor_nf,dict_tables, modalidade,generalidades,cidaten):

    try:
        uf_origem, uf_destino, faixa_preco, faixa_fluvial, faixa_gris = informacoes_cidaten(cep_origem,cep_destino,cidaten)
        frete_peso, kg_adicional = calcular_frete_peso(uf_origem,uf_destino,faixa_preco,peso, modalidade, dict_tables)
        gris = calcular_gris(valor_nf,faixa_gris,generalidades)

    except:
        uf_origem = 'NA'
        uf_destino = 'NA'
        faixa_preco = 'NA'
        faixa_fluvial = 'NA'
        faixa_gris = 'NA'
        frete_peso='NA'
        kg_adicional = 'NA'
        gris= 'NA'



    #print([uf_origem, uf_destino, faixa_preco, faixa_fluvial, faixa_gris, frete_peso, kg_adicional, adv, gris, taxa_nao_mec, fluvial, prod_valioso])
    
    return [uf_origem, uf_destino, faixa_preco, faixa_fluvial, faixa_gris, frete_peso, kg_adicional, gris]

#funcao para trazer informações da faixa de cep:
def informacoes_cidaten(cep_origem,cep_destino,cidaten_utilizar):
    uf_origem = cidaten_utilizar['UF'][(cidaten_utilizar['CEP I']<= cep_origem) & (cidaten_utilizar['CEP F']>= cep_origem)].values.tolist()
    dados_destino = cidaten_utilizar[['UF','Preço','Fluvial','Faixa Gris']][(cidaten_utilizar['CEP I']<= cep_destino) & (cidaten_utilizar['CEP F']>= cep_destino)].iloc[0].tolist()
    return uf_origem + dados_destino

#Funcao para calcular frete peso e kg adicional
def calcular_frete_peso(uf_origem, uf_destino, faixa_preco,peso, modalidade, dict_tables):
    faixa = uf_origem + uf_destino + faixa_preco
    limite_peso = float(dict_tables[modalidade].keys()[-2])
    kg_adicional = 0
    limite_peso
    if peso > limite_peso:
        key = limite_peso
        #tentar achar a chave e calcular o frete peso
        try:
            frete_peso = round(dict_tables[modalidade].loc[faixa,key],2)
            kg_adicional = round(dict_tables[modalidade].loc[faixa,"KG ADICIONAL"]*(math.ceil(peso-limite_peso)),2)
        #caso nao encontre a chave
        except:
            frete_peso = 'SORI'
            kg_adicional = 'SORI'

    else:
        for i in dict_tables[modalidade].keys()[4:-1]: 
            if float(i) >= peso:
                key = i
                #tentar achar a chave e calcular o frete peso
                try:
                    frete_peso = round(dict_tables[modalidade].loc[faixa,key],2)

                #caso não encontre a chave
                except:
                    frete_peso = 'SORI'
                break
    return frete_peso, kg_adicional

#função para calcular o gris
def calcular_gris(valor_nf,faixa_gris,generalidades):
    valor_gris = 0
    if faixa_gris == 1:
        valor_gris = generalidades['gris']['taxa_1']*valor_nf
        if valor_gris < generalidades['gris']['min_1']:
            valor_gris = generalidades['gris']['min_1']
    elif faixa_gris == 2:
        valor_gris = generalidades['gris']['taxa_2']*valor_nf
        if valor_gris < generalidades['gris']['min_2']:
            valor_gris = generalidades['gris']['min_2']
    return round(valor_gris,2)

#função para calcular o advalorem
def calcular_adv(valor_nf,generalidades):
    adv = generalidades['adv']['taxa']*valor_nf
    if adv < generalidades['adv']['min']:
        adv = generalidades['adv']['min']
    return round(adv,2)

#função para calcular o fluvial
def calcular_fluvial(valor_nf, adv, faixa_fluvial, generalidades):
    fluvial = 0
    if faixa_fluvial == 'Fluvial':
        fluvial = generalidades['fluvial']['taxa']*valor_nf
        if fluvial < generalidades['fluvial']['min']:
            fluvial = generalidades['fluvial']['min']
        fluvial -= adv
    return round(fluvial,2)

#funcao para calcular taxa não mecanizável
def calcular_tx_nao_mec(peso, generalidades):
    taxa_nao_mec = 0
    if peso >= generalidades['taxa_nao_mec']['peso_min']:
        taxa_nao_mec = generalidades['taxa_nao_mec']['taxa']
    return round(taxa_nao_mec,2)

def calcular_prod_valioso_cv(valor_nf,peso,adv,faixa_gris,generalidades):
    produto_valioso = 0
    if generalidades['prod_valioso'] == 'cobrar' and valor_nf > 30000:
        produto_valioso = 0.0015*valor_nf
    elif generalidades['prod_valioso'] == 1 and valor_nf/peso > 3000 and valor_nf> 3000:
        produto_valioso = adv
        if faixa_gris > 0:
            produto_valioso += 0.0015 * valor_nf
    
    return round(produto_valioso,2)

def taxa_icms(uf_origem,uf_destino,tabela_icms):
    try:
        taxa = tabela_icms.loc[uf_origem,uf_destino]
    except:
        taxa = 'NA'
    return taxa

#----------------PARAMETROS----------------#

#generalidades a serem utilizadas como parâmetros de cálculos
generalidades = {'adv' : {'taxa': 0.003, 'min': 0.0},
            'gris': {'taxa_1': 0.0034, 'taxa_2': 0.0134, 'min_1': 0.34,'min_2': 1.34},
            'fluvial': {'taxa': 0.08,'min': 8},
            'taxa_nao_mec': {'taxa': 40,'peso_min': 50},
            'prod_valioso': 'cobrar'}

#Caminho da tabela de ICMS
caminho_icms = 'Alíquota de ICMS.xlsx'
tabela_icms = pd.read_excel(caminho_icms,index_col=0, dtype=object)

#Tabela de frete peso a ser utilizada nos calculos
caminho_tabela = 'TEMPLATE_TABELAS_PETLOVE_ATUALIZADO_1608 (1).xlsx'

#Arquivos de cidaten a serem utilizados
caminho_cidaten_padrao = 'cidaten_padrao.xlsx'
caminho_cidaten_interior2 = 'cidaten_interior2.xlsx'

#Caminho dos relatorios
caminho_relatorio_descritivo = 'Descritivo_PETSUPERMARKET.xlsx'
caminho_relatorio_demonstrativo = 'Demonstrativo_PETSUPERMARKET.xlsx'

#ler relatorios
demonstrativo = pd.read_excel(caminho_relatorio_demonstrativo, dtype=object, header=1)
descritivo = pd.read_excel(caminho_relatorio_descritivo, dtype=object, header=1)
print('CARREGAR RELATORIOS - OK')

#Ler tabelas do cliente
tabelas = ['.PACKAGE', '.COM', 'ECONOMICO', 'PICKUP'] #modalidades existentes
tabelas_vazias = [] # lista de tabelas vazias
dict_tables = {}
for tabela in tabelas:

    df_atual = pd.read_excel(caminho_tabela, sheet_name=tabela,dtype=object)
    dict_tables[tabela] = df_atual

    #criar lista de tabelas vazias
    if dict_tables[tabela].empty:
        tabelas_vazias.append(tabela)

    if tabela not in tabelas_vazias:
        uf_destino = dict_tables[tabela]['UF DESTINO'].copy()
        tipo_tarifa = dict_tables[tabela]['TIPO TARIFA'].copy()
        

        dict_tables[tabela]['CHAVE'] = dict_tables[tabela]['UF ORIGEM'].str.cat(uf_destino)
        dict_tables[tabela]['CHAVE'] = dict_tables[tabela]['CHAVE'].str.cat(tipo_tarifa)
        dict_tables[tabela]['CHAVE'] = dict_tables[tabela]['CHAVE'].apply(lambda x: x.replace(' ', ''))
        dict_tables[tabela] = dict_tables[tabela].set_index(['CHAVE'])

print('CARREGAR TABELA CLIENTE - OK')
#Ler cidatens
cidaten_interior2 = pd.read_excel("cidaten_interior2.xlsx",dtype=object)
cidaten_padrao = pd.read_excel("cidaten_padrao.xlsx",dtype=object)
print("CARREGAR CIDATENS - OK")
#Ler tabela de iatas - FUTURO


#------------------------------------------------#

#Pegar colunas relevantes dos relatorios
demonstrativo.drop(['Receb.', 'Frete', 'CPF/CNPJ', 'Peso'], axis=1)
descritivo = descritivo[['Cte', 'Peso Real', 'Peso Taxado', 'Valor', 'Valor Merc', 'Cep Ori', 'Cep Dest', 'Pedido', 'ShipmentId']]

#Unir relatorios (INNER JOIN, EXCLUI O QUE TIVER EM UM E NÃO EM OUTRO)
df_unido = pd.merge(demonstrativo, descritivo, left_on='Remessa',right_on='Cte', how='inner')

#descobrir todas as colunas com componentes
colunas_relatorio = df_unido.columns
colunas_componentes = [coluna for coluna in colunas_relatorio if "COMPONENTES" in coluna]

colunas_nao_calculadas = [coluna for coluna in colunas_componentes if coluna in ['COMPONENTES;TAXA EXTRAORI', 'COMPONENTES;TAXA EXTRADEST', 'COMPONENTES;TAXA_DIF_ENTREGA']]

#Calculando frete taxado
df_unido['FRETE TAXADO'] = df_unido[colunas_componentes].sum(axis=1)

#Definindo colunas a serem utilizadas na macro e ordenando
colunas_macro = ['Cte', 'Operação', 'Data'] + colunas_componentes + ['FRETE TAXADO','Peso Real', 'Peso Taxado', 'Valor Merc', 'Cep Ori', 'Cep Dest', 'Pedido', 'ShipmentId']

#Selecionando apenas as colunas relevantes
df_base = df_unido[colunas_macro]
df_base['Pedido'] = df_base['Pedido'].str.replace(' ', '') #por algum motivo pedido vem cheio de espaço em branco.
print('LIMPAR RELATÓRIOS - OK')
print('REALIZANDO CÁLCULOS - FAVOR AGUARDAR')

#Calcular advalorem
df_base['Ad_valorem'] = generalidades['adv']['taxa']*df_base['Valor Merc']
mask = df_base['Ad_valorem'] < generalidades['adv']['min']
df_base.loc[mask,'Ad_valorem'] = generalidades['adv']['min']
df_base['Ad_valorem'] = pd.to_numeric(df_base['Ad_valorem'])
df_base['Ad_valorem'].round(2)
print('CALCULO ADV - OK')

#Calcular taxa nao mecanizavel
df_base['Taxa_Nao_Mec'] = 0
mask = df_base['Peso Taxado'] > generalidades['taxa_nao_mec']['peso_min']
df_base.loc[mask,'Taxa_Nao_Mec'] = generalidades['taxa_nao_mec']['taxa']
print('CALCULO TAXA_NAO_MEC - OK')

#Calcular produto valioso
if generalidades['prod_valioso'] == 'cobrar':
    df_base['Produto_Valioso'] = np.where(((df_base['Valor Merc']/df_base['Peso Taxado']>3000) & (df_base['Valor Merc']> 3000)), df_base['Ad_valorem'] + 0.0015*df_base['Valor Merc'], 0)
    df_base['Produto_Valioso'] = np.where(((df_base['Valor Merc']/df_base['Peso Taxado']>3000) & (df_base['Valor Merc']> 3000) & (df_base['Valor Merc']> 30000)), 0.0015*df_base['Valor Merc'], 0)
else:
    df_base['Produto_Valioso'] = 0
print('CALCULO PROD VALIOSO - OK')

#Chamar funcao de calculo:
df_base[["uf_origem", "uf_destino", "faixa_preco", "faixa_fluvial", "faixa_gris", "frete_peso", "kg_adicional", "Gris"]] = df_base.apply(lambda x: calcular_componentes(x['Cep Ori'], x['Cep Dest'], x['Peso Taxado'], x['Valor Merc'], dict_tables,x['Operação'], generalidades, cidaten_padrao),axis=1,result_type='expand')
print('FRETE PESO E ADICIONAL - OK')

print('CALCULANDO FLUVIAL')
#Calcular Fluvial
if generalidades['fluvial']['taxa'] > 0:
    df_base['Valor_Fluvial'] = np.where((df_base['faixa_fluvial'] == 'Fluvial'), (df_base['Valor Merc']*generalidades['fluvial']['taxa'] - df_base['Ad_valorem']), 0)
    df_base['Valor_Fluvial'] = np.where(((df_base['Valor_Fluvial'] > 0) & (df_base['Valor_Fluvial'] < generalidades['fluvial']['min'])), (generalidades['fluvial']['min'] - df_base['Ad_valorem']), df_base['Valor_Fluvial'])
else:
    df_base['Valor_Fluvial'] = 0
    
print('BATENDO TABELA ICMS')
df_base['Percentual_ICMS'] = df_base.apply(lambda x: taxa_icms(x['uf_origem'],x['uf_destino'],tabela_icms), axis=1)

print('CALCULANDO FRETE TOTAL')

colunas_recalc_nao_alteradas = []
if len(colunas_nao_calculadas) > 0:
    for coluna in colunas_nao_calculadas:
        df_base['Recalculado_'+coluna[12:]] = df_base[coluna]
        colunas_recalc_nao_alteradas.append('Recalculado_'+coluna[12:])
colunas_frete_total = ['frete_peso', 'kg_adicional', 'Ad_valorem', 'Gris', 'Produto_Valioso', 'Valor_Fluvial'] + colunas_recalc_nao_alteradas

df_base['FRETE TOTAL'] = df_base[colunas_frete_total].replace('[a-zA-Z]', '0', regex=True).astype(float).fillna(0).sum(axis=1)


print('CALCULANDO VALOR FINAL')

df_base['Valor_ICMS'] = df_base['Percentual_ICMS'].replace('[a-zA-Z]', '0', regex=True).astype(float)*(df_base['FRETE TOTAL'].replace('[a-zA-Z]', '0', regex=True).astype(float)/(1-df_base['Percentual_ICMS'].replace('[a-zA-Z]', '0', regex=True).astype(float)))
df_base['FRETE ACORDADO'] = df_base['Valor_ICMS'].replace('[a-zA-Z]', '0', regex=True).astype(float) + df_base['FRETE TOTAL'].replace('[a-zA-Z]', '0', regex=True).astype(float)

df_base[['FRETE TOTAL','Valor_ICMS', 'FRETE ACORDADO']] = df_base[['FRETE TOTAL','Valor_ICMS', 'FRETE ACORDADO']].round(2)

colunas_df_final = colunas_macro + ["uf_origem", "uf_destino", "faixa_preco", "faixa_fluvial","faixa_gris","Percentual_ICMS", "frete_peso", "kg_adicional" ] + ['Ad_valorem', 'Taxa_Nao_Mec', 'Valor_Fluvial','Gris' , 'Produto_Valioso'] + colunas_recalc_nao_alteradas + ['FRETE TOTAL','Valor_ICMS', 'FRETE ACORDADO']

df_final = df_base[colunas_df_final]

df_final.set_index('Cte', inplace=True)
print('SALVANDO EXCEL')
df_final.to_excel('Calculos_realizados.xlsx', header=1)

print(time.time() - time_i)

#print(calcular_componentes('37640000', '58111600', 0.54, 0.64, dict_tables, '.PACKAGE', generalidades, cidaten_padrao))