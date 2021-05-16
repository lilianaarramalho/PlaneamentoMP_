from ct import *
from ov import *
from item import *
from of import *
from cliente import *
import pandas as pd
import time
from functions import *
import math
import numpy as np
import datetime


method=0

print('importar dados')
start=time.time()

def importar_clientes():

    clientes=[]
    df_clientes=pd.read_csv('data/01. clientes.csv',sep=",",encoding='UTF-8')
    df_clientes['Carga Completa']=df_clientes['Carga Completa'].fillna(1)
    df_clientes['Lead Time Acordado (semanas)'] = df_clientes['Lead Time Acordado (semanas)'].fillna(0)
    df_clientes_total=df_clientes.copy()
    df_clientes=df_clientes.drop_duplicates(['Grupo'])
    nome_clientes=df_clientes['Nome'].tolist()
    carga_completa=df_clientes['Carga Completa'].tolist()
    lead_time=df_clientes['Lead Time Acordado (semanas)'].tolist()
    grupo=df_clientes['Grupo'].tolist()

    for index in range(len(nome_clientes)):
        new_cliente=cliente(index,grupo[index],lead_time[index],carga_completa[index])
        clientes.append(new_cliente)

    for index in range(len(clientes)):

        grupo=clientes[index].nome

        df_grupo=df_clientes_total[df_clientes_total['Grupo']==grupo]

        nomes_total = df_grupo['Nome'].tolist()

        for pos_nome in range(len(nomes_total)):

            clientes[index].descricao_clientes.append(nomes_total[pos_nome])

    return clientes

clientes=importar_clientes()

def importar_cts():

    global semana_inicio_plano

    semana_inicio_plano=-1


    cts=[]

    df_cts=pd.read_csv('data/02. cts.csv',sep=",",encoding="UTF-8")
    df_semanas=pd.read_csv('data/03. semanas.csv',sep=",",encoding="UTF-8")
    df_blocos=pd.read_csv('data/04. restricoes blocos.csv',sep=",",encoding="UTF-8")
    df_virados=pd.read_csv('data/05. restricoes virados.csv',sep=",",encoding="UTF-8")
    df_clientes=pd.read_csv('data/06. restricoes clientes.csv',sep=",",encoding='UTF-8')

    df_cts['turnos/dia']=df_cts['turnos/dia'].astype(float)
    df_cts['horas/turno']=df_cts['horas/turno'].astype(float)
    df_cts['OEE']=df_cts['OEE'].astype(float)

    df_cts['capacidade_dia']=df_cts['turnos/dia']*df_cts['horas/turno']*60*df_cts['OEE']
    nome_ct=df_cts['ct'].tolist()
    acabamento=df_cts['acabamento'].tolist()
    dias_semana=df_semanas['dias'].tolist()
    current_week=datetime.datetime.now().isocalendar()[1]
    posicoes_a_remover=current_week-1
    current_day=datetime.datetime.now().weekday()
    if current_day>4:
        dias_semana = dias_semana[posicoes_a_remover+1:]
        semana_inicio_plano=current_week+1
    else:
        dias_semana = dias_semana[posicoes_a_remover + 1:]
        n_dias_passados=current_day+1
        if n_dias_passados >=dias_semana[0]:
            dias_semana=dias_semana[1:]
        else:
            dias_semana[0]=dias_semana[0]-n_dias_passados

    capacidade_dia=df_cts['capacidade_dia'].tolist()

    ct_blocos=df_blocos['ct'].tolist()
    capacidade_blocos=df_blocos['blocos semana'].tolist()

    ct_virados=df_virados['ct'].tolist()
    capacidade_virados=df_virados['blocos semana'].tolist()

    ct_clientes=df_clientes['ct'].tolist()
    grupo_cliente=df_clientes['cliente'].tolist()
    capacidade_cliente=df_clientes['capacidade'].tolist()

    for id_ct in range(len(nome_ct)):

        capacidade_bl=0
        capacidade_vir=0

        new_ct=ct(id_ct,nome_ct[id_ct],acabamento[id_ct])
        cts.append(new_ct)

        if nome_ct[id_ct] in ct_blocos:
            posicao_blocos = ct_blocos.index(nome_ct[id_ct])
            capacidade_bl = capacidade_blocos[posicao_blocos]

        if nome_ct[id_ct] in ct_virados:
            posicao_virados = ct_virados.index(nome_ct[id_ct])
            capacidade_vir = capacidade_virados[posicao_virados]

        for index in range(len(dias_semana)):
            new_ct.capacidade.append(int(dias_semana[index])*int(capacidade_dia[id_ct]))
            new_ct.capacidade_virados.append(capacidade_vir)
            new_ct.capacidade_blocos.append(capacidade_bl)
            new_ct.capacidade_iniciais.append(dias_semana[index]*capacidade_dia[id_ct])

        for posicao in range(len(clientes)):
            capacidade_cliente_atual = [i * 0 for i in new_ct.capacidade_iniciais]
            new_ct.capacidade_clientes.append(capacidade_cliente_atual)

        if nome_ct[id_ct] in ct_clientes:

            posicao_clientes=find(ct_clientes,nome_ct[id_ct])

            for id_posicao in range(len(posicao_clientes)):

                posicao=posicao_clientes[id_posicao]

                grupo_tofind=grupo_cliente[posicao]
                capacidade=capacidade_cliente[posicao]

                for id_cliente in range(len(clientes)):

                    if clientes[id_cliente].nome==grupo_tofind:

                        capacidade_cliente_atual = [i * capacidade for i in new_ct.capacidade_iniciais]

                        new_ct.capacidade_clientes[id_cliente]=capacidade_cliente_atual
    return cts

cts=importar_cts()

def import_ofs(method):

    ofs=[]


    df_ofs = pd.read_csv('data/08.  ofs.csv', sep=",", encoding='UTF-8')

    df_ofs=df_ofs.drop_duplicates()
    df_ofs = df_ofs[df_ofs['Ordem de produção / planeada'].notnull()]
    df_ofs['Ordem de produção / planeada']=df_ofs['Ordem de produção / planeada'].astype(int)
    df_ofs['Qtd.necessária'] = df_ofs['Qtd.necessária'].str.replace(' ', '')
    df_ofs['Qtd.necessária'] = df_ofs['Qtd.necessária'].astype(float)
    df_ofs['Duração da operação'] = df_ofs['Duração da operação'].str.replace(' ', '')
    df_ofs['Duração da operação'] = df_ofs['Duração da operação'].astype(float)

    if method == 0:  # por planear
        df_ofs=df_ofs[(df_ofs['Ordem de produção / planeada']<400000000) & (df_ofs['Ordem de produção / planeada']>300000000)]
    elif method==1: # só não planeadas
        df_ofs = df_ofs[(df_ofs['Ordem de produção / planeada'] > 1600000000)]

    df_ofs['planeador']=df_ofs['Nome planejador'].str.split(' ').str[0]
    df_ofs['acabamento']=df_ofs['Texto breve de material'].str.split('/').str[1]
    df_ofs['acabamento'] = df_ofs['acabamento'].str.split(' ').str[0]
    df_ofs['descricao_bloco']=df_ofs['Descritivo componente'].str.split(' ').str[0]
    df_ofs = df_ofs.sort_values(['Ordem Venda / Transferência', 'Item OV/Transferência', 'Sold to'], ascending=[False, False, False])
    df_ofs['Data desejada de saída de mercadoria'] = df_ofs['Data desejada de saída de mercadoria'].fillna(
        datetime.datetime.now())
    df_ofs['Data desejada de saída de mercadoria'] = df_ofs['Data desejada de saída de mercadoria'].astype(
        'datetime64[ns]')
    df_ofs['Data-base do início']=df_ofs['Data-base do início'].fillna(
        datetime.datetime.now())
    df_ofs['Data-base do início'] = df_ofs['Data-base do início'].astype(
        'datetime64[ns]')

    lista_ovs=df_ofs['Ordem Venda / Transferência'].tolist()
    cod_ofs = df_ofs['Ordem de produção / planeada'].tolist()
    lista_items = df_ofs['Item OV/Transferência'].tolist()

    cod_ov_in=1000

    for index in range(len(cod_ofs)):
        if math.isnan(lista_ovs[index]):
            lista_ovs[index]=cod_ov_in
            cod_ov_in+=1

        if math.isnan(lista_items[index]):
            lista_items[index]=99

    df_ofs['ov'] = lista_ovs
    df_ofs['item']=lista_items
    df_ofs=df_ofs.sort_values(['ov', 'item'], ascending=[False, False])
    df_cliente_final=df_ofs.copy()
    df_cliente_final=df_cliente_final[['ov', 'item','Sold to','planeador','Data desejada de saída de mercadoria']]
    df_cliente_final=df_cliente_final[df_cliente_final['Sold to'].notnull()]
    df_cliente_final['key']=df_cliente_final['ov'] + df_cliente_final['item']
    df_cliente_final=df_cliente_final[['key','Sold to', 'planeador','Data desejada de saída de mercadoria']]
    df_cliente_final=df_cliente_final.sort_values(['key','planeador'],ascending=True)
    df_cliente_final=df_cliente_final.drop_duplicates(['key'])
    df_cliente_final['Data desejada de saída de mercadoria'] = df_cliente_final[
        'Data desejada de saída de mercadoria'].fillna(datetime.datetime.now())
    df_ofs['key'] = df_ofs['ov'] + df_ofs['item']


    df_ofs=df_ofs.merge(df_cliente_final,how='left',on='key')
    df_ofs['Data desejada de saída de mercadoria_y'] = df_ofs[
        'Data desejada de saída de mercadoria_y'].fillna(datetime.datetime.now())
    df_ofs['semana'] = df_ofs.apply(lambda row: row['Data desejada de saída de mercadoria_y'].isocalendar()[1], axis=1)
    df_ofs = df_ofs.rename(columns={'Data desejada de saída de mercadoria_y': 'Data desejada de saída de mercadoria'})
    df_ovs=df_ofs.copy()

    df_ovs = df_ovs.drop_duplicates(['ov','semana'])

    ovs_ov=df_ovs['ov'].tolist()
    id_cliente_interno=df_ovs['planeador_y'].tolist()
    id_cliente = df_ovs['Sold to_y'].tolist()
    data_desejada = df_ovs['Data desejada de saída de mercadoria'].tolist()

    ovs = []

    for index in range(len(ovs_ov)):

        semana=data_desejada[index].isocalendar()[1]
        cod_ov=ovs_ov[index]
        cliente=id_cliente[index]
        cliente_interno=id_cliente_interno[index]

        id_cliente_ov=-1
        id_interno_ov=-1

        for id_lista in range(len(clientes)):
            lista_clientes=clientes[id_lista].descricao_clientes
            if cliente in lista_clientes:
                id_cliente_ov=id_lista
                break

        if cliente_interno=='CCS':
            for id_lista in range(len(clientes)):
                if clientes[id_lista].nome=='cliente interno':
                    id_interno_ov = id_lista
                    break

        new_ov=ov(index,cod_ov,id_cliente_ov,id_interno_ov,semana)
        ovs.append(new_ov)

    df_items = df_ofs.copy()
    df_items = df_items.drop_duplicates(['ov','item','semana'])
    items=[]

    ovs_item = df_items['ov'].tolist()
    items_item = df_items['item'].tolist()

    for index in range(len(ovs_item)):

        cod_item=items_item[index]

        for pos_ov in range(len(ovs)):
            if ovs_item[index]==ovs[pos_ov].cod_ov:
                id_ov=pos_ov
                ovs[pos_ov].id_items.append(index)

        new_item=item(index,cod_item,id_ov)
        items.append(new_item)

    ops=df_ofs['Ordem de produção / planeada'].tolist()
    duracoes = df_ofs['Duração da operação'].tolist() * 60
    quantidades = df_ofs['Quantidade total da ordem'].tolist()
    codigos_material = df_ofs['Material de produção'].tolist()
    descricoes_material = df_ofs['Texto breve de material'].tolist()
    codigos_precedencia = df_ofs['Componente'].tolist()
    blocos = df_ofs['Qtd.necessária'].tolist()
    descricao_bloco = df_ofs['descricao_bloco'].tolist()
    data_alocada = df_ofs['Data-base do início'].tolist()
    data_desejada = df_ofs['Data desejada de saída de mercadoria'].tolist()
    lista_items=df_ofs['item'].tolist()
    lista_ovs=df_ofs['ov'].tolist()
    df_acabamentos = pd.read_csv('data/07.  acabamentos.csv', sep=",", encoding='UTF-8')
    lista_descricao_acabamento=df_acabamentos['acabamento'].tolist()
    lista_tipo_acabamento=df_acabamentos['tipo'].tolist()
    lista_cts=df_ofs['Centro de trabalho'].tolist()
    lista_acabamentos=df_ofs['acabamento'].tolist()


    ofs=[]
    for index in range(len(ops)):

        cod_of=ops[index]
        duracao=duracoes[index]
        quantidade=quantidades[index]
        codigo_material=codigos_material[index]
        descricao_material=descricoes_material[index]
        viradas=0 #TODO adicionar viradas
        codigo_precedencia=codigos_precedencia[index]

        if method == 0:  # por planear
            id_alocada = -1
        elif method == 1:  # só não planeadas
            id_alocada=data_alocada[index].isocalendar()[1]
            df_ofs = df_ofs[(df_ofs['Ordem de produção / planeada'] > 1600000000)]

        if descricao_bloco[index]=='BL':
            n_blocos=int(float(blocos[index]))
        else:
            n_blocos=0

        if lista_acabamentos[index] in lista_descricao_acabamento:
            posicao=lista_descricao_acabamento.index(lista_acabamentos[index])
            acabamento=lista_tipo_acabamento[posicao]
        else:
            acabamento='Qualquer'

        id_ct=-1

        for pos_ct in range(len(cts)):

            if cts[pos_ct].nome==lista_cts[index] and cts[pos_ct].acabamento==acabamento:

                id_ct=pos_ct

                break

        id_items=[]

        for pos_item in range(len(items)):

            id_ov=items[pos_item].id_ov

            if lista_items[index]==items[pos_item].cod_item and lista_ovs[index]==ovs[id_ov].cod_ov and data_desejada[index]==ovs[id_ov].data_desejada:

                id_items.append(pos_item)

        new_of = of(index, cod_of, duracao, quantidade, id_ct, codigo_material, descricao_material, codigo_precedencia,
                    id_items, blocos, viradas)

        new_of.id_alocada=id_alocada

        ofs.append(new_of)

    return ovs,items,ofs

ovs,items,ofs=import_ofs(method)

def atualizar_capacidades(method):

    if method==0:

        df_ofs = pd.read_csv('data/08.  ofs.csv', sep=",", encoding='UTF-8')
        df_ofs = df_ofs.drop_duplicates()
        df_ofs = df_ofs[df_ofs['Ordem de produção / planeada'].notnull()]
        df_ofs['Ordem de produção / planeada'] = df_ofs['Ordem de produção / planeada'].astype(int)
        df_ofs = df_ofs[(df_ofs['Ordem de produção / planeada'] > 1600000000)]
        df_ofs['Data-base do fim'] = df_ofs['Data-base do fim'].fillna(
            datetime.datetime.now())
        df_ofs['Data-base do fim'] = df_ofs['Data-base do fim'].astype(
            'datetime64[ns]')
        df_ofs['semana'] = df_ofs.apply(lambda row: row['Data-base do fim'].isocalendar()[1],axis=1)
        df_ofs['Duração da operação'] = df_ofs['Duração da operação'].str.replace(' ', '')
        df_ofs['Duração da operação'] = df_ofs['Duração da operação'].astype(float)

        #TODO ADICIONAR RESTRIÇAO CAPACIDADE CLIENTES,BLOCOS E VIRADAS

        df_ofs=df_ofs[df_ofs['semana']>=semana_inicio_plano]

        df_ofs['semana']=df_ofs['semana']-semana_inicio_plano

        df_ofs=df_ofs[['Centro de trabalho','semana','Duração da operação']]

        df_ofs=df_ofs.groupby(by=['Centro de trabalho','semana'])['Duração da operação'].sum()
        df_ofs=df_ofs.reset_index()

        lista_cts=df_ofs['Centro de trabalho'].tolist()
        lista_semanas=df_ofs['semana'].tolist()
        lista_duracao=df_ofs['Duração da operação'].tolist()

        for index in range(len(lista_cts)):

            for id_ct in range(len(cts)):

                if cts[id_ct].nome==lista_cts[index]:

                    pos_semana=lista_semanas[index]

                    cts[id_ct].capacidade[pos_semana]=cts[id_ct].capacidade[pos_semana]-lista_duracao[index]*60


atualizar_capacidades(method)

end=time.time()
print('tempo: ' + str(end-start))
