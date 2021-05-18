import bisect
from ct import *
from ov import *
from item import *
from of import *
from cliente import *
import pandas as pd
import time
import math
import numpy as np
import datetime
from fractions import Fraction
from tqdm import tqdm
import time


method=0

def find(lst, a):
    result = []
    for i, x in enumerate(lst):
        if x==a:
            result.append(i)
    return result


print('importar clientes')
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

    with tqdm(total=len(clientes)) as pbar:

        for index in range(len(clientes)):

            grupo=clientes[index].nome

            df_grupo=df_clientes_total[df_clientes_total['Grupo']==grupo]

            nomes_total = df_grupo['Nome'].tolist()

            for pos_nome in range(len(nomes_total)):

                clientes[index].descricao_clientes.append(nomes_total[pos_nome])

            pbar.update(1)

    pbar.close()

    return clientes

clientes=importar_clientes()

print('importar centros de trabalho')

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
        semana_inicio_plano = current_week
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

    with tqdm(total=len(nome_ct)) as pbar:

        for id_ct in range(len(nome_ct)):

            capacidade_bl=9999999999
            capacidade_vir=9999999999

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
                capacidade_cliente_atual = [i * 99999999 for i in new_ct.capacidade_iniciais]
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

            pbar.update(1)

    pbar.close()

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

    df_ofs['planeador']=df_ofs.iloc[:,0].str.split(' ').str[0]
    df_ofs['acabamento']=df_ofs['Texto breve de material'].str.split('/').str[1]
    df_ofs['acabamento'] = df_ofs['acabamento'].str.split(' ').str[0]
    df_ofs['descricao_bloco']=df_ofs['Descritivo componente'].str.split(' ').str[0]
    df_ofs = df_ofs.sort_values(['Ordem Venda / Transferência', 'Item OV/Transferência', 'Sold to'], ascending=[False, False, False])

    df_ofs['Data-base do fim'] = pd.to_datetime(df_ofs['Data-base do fim'], format='%d/%m/%Y')

    df_ofs['Data-base do início']=df_ofs['Data-base do início'].fillna(datetime.datetime.now())

    df_ofs['Data de criação da ov'] = pd.to_datetime(df_ofs['Data de criação da ov'],format='%d/%m/%Y')

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

    df_ofs=df_ofs.sort_values(['ov', 'item','Material de produção'], ascending=[False, False,False])

    df_ofs['key'] = df_ofs['ov'].astype(str) + ' ' + df_ofs['item'].astype(str)

    df_cliente_final=df_ofs.copy()
    df_cliente_final=df_cliente_final[['ov', 'item','Sold to','planeador','Data-base do fim','Data de criação da ov']]
    #df_cliente_final=df_cliente_final[df_cliente_final['Sold to'].notnull()] como existem ovs que não têm sold to, é melhor remover duplicados e ficar com o primeiro
    df_cliente_final['key']=df_cliente_final['ov'].astype(str) + ' ' + df_cliente_final['item'].astype(str)
    df_cliente_final = df_cliente_final.drop_duplicates(subset=['key'], keep='first')
    df_cliente_final=df_cliente_final[['key','Sold to', 'planeador','Data-base do fim','Data de criação da ov']]
    df_cliente_final=df_cliente_final.sort_values(['key','planeador'],ascending=True)
    df_cliente_final=df_cliente_final.drop_duplicates(['key'])


    df_cliente_final['Data-base do fim'] = df_cliente_final[
        'Data-base do fim'].fillna(datetime.datetime.now())

    df_cliente_final['Data de criação da ov'] = df_cliente_final[
        'Data de criação da ov'].fillna(datetime.datetime.now())


    df_ofs=df_ofs.merge(df_cliente_final,how='left',on='key')
    df_ofs['semana'] = df_ofs.apply(lambda row: row['Data-base do fim_y'].isocalendar()[1], axis=1)
    df_ofs['semana']=df_ofs['semana']-semana_inicio_plano-1
    df_ofs = df_ofs.rename(columns={'Data-base do fim_y': 'Data-base do fim'})
    df_ofs = df_ofs.rename(columns={'Data de criação da ov_y': 'Data de criação da ov'})
    df_ofs['Data de criação da ov'] = df_ofs[
        'Data de criação da ov'].fillna(datetime.datetime.now())
    df_ofs['semana_criacao'] = df_ofs.apply(lambda row: row['Data de criação da ov'].isocalendar()[1], axis=1)
    df_ofs['semana_criacao']=df_ofs['semana_criacao']-semana_inicio_plano+1

    df_ovs=df_ofs.copy()

    df_ovs = df_ovs.drop_duplicates(['ov','semana'])

    ovs_ov=df_ovs['ov'].tolist()
    id_cliente_interno=df_ovs['planeador_y'].tolist()
    id_cliente = df_ovs['Sold to_y'].tolist()
    data_desejada = df_ovs['semana'].tolist()
    semanas_criacao=df_ovs['semana_criacao'].tolist()

    ovs = []

    print('importar ordens de venda')

    with tqdm(total=len(ovs_ov)) as pbar:

        for index in range(len(ovs_ov)):

            semana=data_desejada[index]
            cod_ov=ovs_ov[index]
            if cod_ov==3826000005:
                print('x')
            cliente=id_cliente[index]
            cliente_interno=id_cliente_interno[index]
            semana_criacao=semanas_criacao[index]

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

            new_ov=ov(index,cod_ov,id_cliente_ov,id_interno_ov,semana,semana_criacao)
            ovs.append(new_ov)
            pbar.update(1)

    pbar.close()

    df_items = df_ofs.copy()
    df_items = df_items.drop_duplicates(['ov','item','semana'])
    items=[]

    ovs_item = df_items['ov'].tolist()
    items_item = df_items['item'].tolist()
    semana_item=df_items['semana'].tolist()

    for index in range(len(ovs_item)):

        cod_item=items_item[index]

        for pos_ov in range(len(ovs)):
            if ovs_item[index]==ovs[pos_ov].cod_ov and semana_item[index]==ovs[pos_ov].data_desejada:
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
    data_desejada = df_ofs['semana'].tolist()
    lista_items=df_ofs['item'].tolist()
    lista_ovs=df_ofs['ov'].tolist()
    df_acabamentos = pd.read_csv('data/07.  acabamentos.csv', sep=",", encoding='UTF-8')
    lista_descricao_acabamento=df_acabamentos['acabamento'].tolist()
    lista_tipo_acabamento=df_acabamentos['tipo'].tolist()
    lista_cts=df_ofs['Centro de trabalho'].tolist()
    lista_acabamentos=df_ofs['acabamento'].tolist()

    pbar.close()

    ofs=[]

    print('importar ordens de fabrico')

    with tqdm(total=len(ops)) as pbar:

        for index in range(len(ops)):

            cod_of=ops[index]
            duracao=duracoes[index]
            quantidade=quantidades[index]
            codigo_material=codigos_material[index]
            descricao_material=descricoes_material[index]
            viradas=0 #TODO adicionar viradas
            codigo_precedencia=codigos_precedencia[index]

            id_alocada = -1

            if method == 0:  # por planear
                id_alocada = -1
            # elif method == 1:  # só não planeadas
            #     id_alocada=data_alocada[index]-semana_inicio_plano
            #     df_ofs = df_ofs[(df_ofs['Ordem de produção / planeada'] > 1600000000)]
            # elif method==2:
            #     id_alocada=data_alocada[index]-semana_inicio_plano

            if descricao_bloco[index]=='BL':
                n_blocos=float(blocos[index])
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
                    items[pos_item].id_ofs.append(index)


            new_of = of(index, cod_of, duracao*60, quantidade, id_ct, codigo_material, descricao_material, codigo_precedencia,
                        id_items, n_blocos, viradas)

            new_of.id_alocada=id_alocada

            ofs.append(new_of)

            pbar.update(1)

    pbar.close()

    return ovs,items,ofs

ovs,items,ofs=import_ofs(method)

print('importar capacidade ocupada')

def atualizar_capacidades(method):



    if method==0:

        df_ofs = pd.read_csv('data/08.  ofs.csv', sep=",", encoding='UTF-8')
        df_acabamentos = pd.read_csv('data/07.  acabamentos.csv', sep=",", encoding='UTF-8')
        lista_descricao_acabamento = df_acabamentos['acabamento'].tolist()
        lista_tipo_acabamento = df_acabamentos['tipo'].tolist()
        df_ofs = df_ofs.drop_duplicates()
        df_ofs = df_ofs[df_ofs['Ordem de produção / planeada'].notnull()]
        df_ofs['Ordem de produção / planeada'] = df_ofs['Ordem de produção / planeada'].astype(int)
        df_ofs = df_ofs[(df_ofs['Ordem de produção / planeada'] > 1600000000)]
        df_ofs['Data-base do fim'] = df_ofs['Data-base do fim'].fillna(
            datetime.datetime.now())
        df_ofs['acabamento'] = df_ofs['Texto breve de material'].str.split('/').str[1]
        df_ofs['acabamento'] = df_ofs['acabamento'].str.split(' ').str[0]

        df_ofs['Data-base do fim'] = pd.to_datetime(df_ofs['Data-base do fim'], format='%d/%m/%Y')

        df_ofs['semana'] = df_ofs.apply(lambda row: row['Data-base do fim'].isocalendar()[1],axis=1)
        df_ofs['Duração da operação'] = df_ofs['Duração da operação'].str.replace(' ', '')
        df_ofs['Duração da operação'] = df_ofs['Duração da operação'].astype(float)

        #TODO ADICIONAR RESTRIÇAO CAPACIDADE CLIENTES,BLOCOS E VIRADAS

        df_ofs=df_ofs[df_ofs['semana']>=semana_inicio_plano]

        df_ofs['semana']=df_ofs['semana']-semana_inicio_plano

        df_ofs=df_ofs[['Centro de trabalho','semana','Duração da operação','acabamento']]

        lista_cts=df_ofs['Centro de trabalho'].tolist()
        lista_semanas=df_ofs['semana'].tolist()
        lista_duracao=df_ofs['Duração da operação'].tolist()
        lista_acabamentos=df_ofs['acabamento'].tolist()

        with tqdm(total=len(lista_cts)) as pbar:

            for index in range(len(lista_cts)):
                if lista_acabamentos[index] in lista_descricao_acabamento:
                    posicao = lista_descricao_acabamento.index(lista_acabamentos[index])
                    acabamento = lista_tipo_acabamento[posicao]
                else:
                    acabamento = 'Qualquer'

                for id_ct in range(len(cts)):

                    if cts[id_ct].nome==lista_cts[index] and cts[id_ct].acabamento==acabamento:

                        pos_semana=lista_semanas[index]

                        cts[id_ct].capacidade[pos_semana]=cts[id_ct].capacidade[pos_semana]-lista_duracao[index]*60

                pbar.update(1)
    pbar.close()


atualizar_capacidades(method)


end=time.time()
#print('tempo: ' + str(end-start))

def print_capacidade_reservada():

    rows=[]

    for index in range(len(cts)):

        for id_ct in range(len(cts)):

            for semana in range(len(cts[id_ct].capacidade)):
                new_row = {'centro de trabalho': cts[id_ct].nome, 'acabamento': cts[id_ct].acabamento,
                           'semana': semana + semana_inicio_plano+1,
                           'carga ocupada': cts[id_ct].capacidade_iniciais[semana] - cts[id_ct].capacidade[semana]}

                rows.append(new_row)

    df = pd.DataFrame(rows)
    df.to_csv('data/11. ocupacao ofs planeadas.csv')




def sort_by_leadtime():

    global ovs
    global clientes

    sorted_index=[]
    sorted_leadtimes=[]

    for index in range(len(ovs)):

        id_cliente=int(ovs[index].id_cliente)

        leadtime_cliente=clientes[id_cliente].lead_time

        semana_desejada=ovs[index].data_desejada

        semana_criacao=ovs[index].data_criacao

        if leadtime_cliente!=0:

            data_com_leadtime = semana_criacao + leadtime_cliente

            data_min=min(semana_desejada,data_com_leadtime)

            ovs[index].data_min=data_min

        else:

            data_min=semana_desejada
        posicao_alocar = bisect.bisect_left(sorted_leadtimes, data_min)
        bisect.insort(sorted_leadtimes, data_min)
        sorted_index.insert(posicao_alocar,index)

    return sorted_index

def verificar_capacidades(id_of,int_semana,id_ov,capacidade_max):

    global ofs
    global cts
    global ovs

    id_ct = ofs[id_of].id_ct

    resultado=True

    #verificar blocos
    blocos=ofs[id_of].blocos
    viradas = ofs[id_of].viradas
    duracao = int(ofs[id_of].duracao)
    cliente_interno = ovs[id_ov].id_interno
    cliente_final = ovs[id_ov].id_cliente

    if cts[id_ct].capacidade_blocos[int_semana]<blocos:
        resultado=False

    #verificar viradas

    if cts[id_ct].capacidade_virados[int_semana]<viradas:
        resultado=False

    #verificar cliente interno

    if cliente_interno!=-1:
        if cts[id_ct].capacidade_clientes[cliente_interno][int_semana]<duracao:
            resultado=False

    # verificar cliente final

    if cliente_final!=-1:
        if cts[id_ct].capacidade_clientes[cliente_final][int_semana]<duracao:
            resultado=False

    # verificar capacidade

    if cts[id_ct].capacidade[int_semana] - duracao < (1-capacidade_max)*cts[id_ct].capacidade_iniciais[int_semana]:
        resultado=False

    return resultado

def alocar(id_of,int_semana,id_ov):

    global ofs
    global cts
    global ovs

    id_ct = ofs[id_of].id_ct
    blocos = ofs[id_of].blocos
    viradas = ofs[id_of].viradas
    duracao = int(ofs[id_of].duracao)
    cliente_interno = ovs[id_ov].id_interno
    cliente_final = ovs[id_ov].id_cliente


    cts[id_ct].capacidade_blocos[int_semana]=cts[id_ct].capacidade_blocos[int_semana]-blocos
    cts[id_ct].capacidade_virados[int_semana]=cts[id_ct].capacidade_virados[int_semana]-viradas
    if cliente_interno!=-1:
        cts[id_ct].capacidade_clientes[cliente_interno][int_semana] = cts[id_ct].capacidade_clientes[cliente_interno][int_semana]-duracao
    if cliente_final!=-1:
        cts[id_ct].capacidade_clientes[cliente_final][int_semana] = cts[id_ct].capacidade_clientes[cliente_final][int_semana] - duracao
    cts[id_ct].capacidade[int_semana]=cts[id_ct].capacidade[int_semana]-duracao

    ofs[id_of].id_alocada=int_semana

def desalocar(id_of,int_semana,id_ov):

    global ofs
    global cts
    global ovs

    id_ct = ofs[id_of].id_ct
    blocos = ofs[id_of].blocos
    viradas = ofs[id_of].viradas
    duracao = int(ofs[id_of].duracao)
    cliente_interno = ovs[id_ov].id_interno
    cliente_final = ovs[id_ov].id_cliente

    cts[id_ct].capacidade_blocos[int_semana] = cts[id_ct].capacidade_blocos[int_semana] + blocos
    cts[id_ct].capacidade_virados[int_semana] = cts[id_ct].capacidade_virados[int_semana] + viradas
    if cliente_interno != -1:
        cts[id_ct].capacidade_clientes[cliente_interno][int_semana] = cts[id_ct].capacidade_clientes[cliente_interno][int_semana] + duracao
    if cliente_final != -1:
        cts[id_ct].capacidade_clientes[cliente_final][int_semana] = cts[id_ct].capacidade_clientes[cliente_final][
                                                                        int_semana] + duracao
    cts[id_ct].capacidade[int_semana] = cts[id_ct].capacidade[int_semana] + duracao

    ofs[id_of].id_alocada = -1

def verificar_precedencias():

    global ofs
    global items

    for index in range(len(items)):

        id_item=items[index].id

        lista_ofs=items[id_item].id_ofs

        codigos_precedencias=[]
        codigos_materiais=[]

        for pos_lista in range(len(lista_ofs)):

            id_of=lista_ofs[pos_lista]

            codigos_precedencias.append(ofs[id_of].codigo_precedencia)
            codigos_materiais.append(ofs[id_of].codigo_material)

        for pos_lista in range(len(lista_ofs)):

            id_of=lista_ofs[pos_lista]

            if ofs[id_of].codigo_precedencia in codigos_materiais:

                lista_posicoes=find(codigos_materiais,ofs[id_of].codigo_precedencia)

                for pos in range(len(lista_posicoes)):

                    pos_lista_ofs=lista_posicoes[pos]

                    id_of_precedencia=lista_ofs[pos_lista_ofs]

                    ofs[id_of].precedencias.append(id_of_precedencia)
                    ofs[id_of_precedencia].sequencias.append(id_of)

def verificar_possivel_atras(id_of,semana_pretendida,id_ov):

    global ofs
    global cts
    global ovs

    possivel=False
    count=semana_pretendida-1
    semana_possivel=-1

    while count>=0 and possivel==False:
        possivel=verificar_capacidades(id_of,count,id_ov,0.95)
        if possivel == True:
            semana_possivel=count
        count-=1

    return semana_possivel

def minimizar_wip_item(id_item,max_semana):

    lista_ofs = items[id_item].id_ofs
    id_ov=items[id_item].id_ov

    if max_semana==-1: #aqui quero minimizar dentro do mesmo item, por isso tenho de calcular a max_semana real

        for index in range(len(lista_ofs)):

            id_of=lista_ofs[index]

            semana_alocada=ofs[id_of].id_alocada

            if semana_alocada>max_semana:
                max_semana=semana_alocada

    for i in range(len(lista_ofs)-1,0-1,-1):

        id_of=lista_ofs[i]

        of=ofs[id_of]

        semana_min_precedencia=0

        for pos_precedencia in range(len(of.precedencias)):
            id_precedencia=ofs[of.precedencias[pos_precedencia]].id
            if ofs[id_precedencia].id_alocada>semana_min_precedencia:
                semana_min_precedencia=ofs[id_precedencia].id_alocada

        if ofs[id_of].id_alocada!=max_semana:
            count = max_semana
            limite=semana_min_precedencia
            duracao=ofs[id_of].duracao

            alterada = False

            while alterada==False and count>=limite:

                id_ct=ofs[id_of].id_ct

                if verificar_capacidades(id_of,count,id_ov,0.95)==True:
                    semana_anterior=ofs[id_of].id_alocada
                    if id_ct!=-1:
                        desalocar(id_of,ofs[id_of].id_alocada,id_ov)
                        alocar(id_of,count,id_ov)
                    alterada=True

                count-=1

            if len(ofs[id_of].precedencias)>0:
                id_precedencia=ofs[id_of].precedencias[0]
                if ofs[id_of].id_alocada<ofs[id_precedencia].id_alocada:
                    alocada=False
                    desalocar(id_precedencia, ofs[id_precedencia].id_alocada, id_ov)
                    for index in range( ofs[id_of].id_alocada,0-1,-1):
                        if verificar_capacidades(id_precedencia,index,id_ov,0.95)==True:
                            alocar(id_precedencia, index, id_ov)
                            alocada=True
                            break
            #         if alocada==False:
            #             semana_impossivel = True
            #             id_ct = ofs[id_of].id_ct
            #             semana_pretendida=ofs[id_of].id_alocada
            #             n_semanas = len(cts[id_ct].capacidade) - semana_pretendida - 1
            #             while semana_impossivel == True and semana_pretendida < n_semanas:
            #                 semana_pretendida += 1
            #                 if verificar_capacidades(id_of, semana_pretendida, id_ov) == True and verificar_capacidades(id_precedencia, semana_pretendida, id_ov) == True:
            #                     semana_impossivel = False
            #                     desalocar(id_of,ofs[id_of].id_alocada,id_ov)
            #                     desalocar(id_precedencia, ofs[id_precedencia].id_alocada, id_ov)
            #                     alocar(id_of, semana_pretendida, id_ov)
            #                     alocar(id_precedencia, semana_pretendida, id_ov)


def minimizar_wip(id_ov):

    max_semana=-1

    for index in range(len(ovs[id_ov].id_items)):
        n_ofs=len(items[ovs[id_ov].id_items[index]].id_ofs)
        for i in range(len(items[ovs[id_ov].id_items[index]].id_ofs)):
            semana_alocada=ofs[items[ovs[id_ov].id_items[index]].id_ofs[i]].id_alocada
            if semana_alocada>max_semana:
                max_semana=semana_alocada

    id_items=ovs[id_ov].id_items

    for index in range(len(id_items)):

        id_item=id_items[index]
        minimizar_wip_item(id_item,max_semana)

    return 0

def print_output():

    global semana_inicio_plano

    rows=[]

    for index in range(len(ofs)):

        new_row={'of':ofs[index].cod_of,
                 'item':items[ofs[index].id_items[0]].cod_item,
                 'id ov':items[ofs[index].id_items[0]].id_ov,
                 'ov':ovs[items[ofs[index].id_items[0]].id_ov].cod_ov,
                 'semana':ofs[index].id_alocada + semana_inicio_plano+1,
                 'data desejada':ovs[items[ofs[index].id_items[0]].id_ov].data_desejada+semana_inicio_plano+1,
                 'centro de trabalho':cts[ofs[index].id_ct].nome,
                 'acabamento': cts[ofs[index].id_ct].acabamento,
                 'blocos':ofs[index].blocos,
                 'ccs':ovs[items[ofs[index].id_items[0]].id_ov].id_interno,
                 'virados':ofs[index].viradas}

        rows.append(new_row)

    df = pd.DataFrame(rows)
    if method==0:
        df.to_csv('data/09. output.csv')
    elif method==1:
        df.to_csv('data/12. output planeado.csv')


def print_mapa_precedencias():

    rows=[]

    for index in range(len(ofs)):

        lista_precedencias=ofs[index].precedencias

        lista_semanas=[]

        max_precedencia=-1

        for pos_precedencia in range(len(lista_precedencias)):

            id_precedencia=lista_precedencias[pos_precedencia]

            semana_precedencia=ofs[id_precedencia].id_alocada

            lista_semanas.append(semana_precedencia)

            max_precedencia=max(lista_semanas)

        new_row = {'of': ofs[index].cod_of, 'item': items[ofs[index].id_items[0]].cod_item,
                   'ov': ovs[items[ofs[index].id_items[0]].id_ov].cod_ov, 'semana': ofs[index].id_alocada,
                   'semana precedencia': max_precedencia}

        rows.append(new_row)

    df = pd.DataFrame(rows)
    df.to_csv('data/10. mapa precedencias.csv')
















