
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

ler_cts_embalagem()

import_bom()

calcular_n_ofs(0.5)

parametro_duracao=60

verificar_precedencias()

id_ovs=sort_by_leadtime()

print_capacidade_reservada(method)

atualizar_capacidade_ref_c(method)

print_capacidades()

count=0

len_in=len(ofs)

start = time.time()

importar_capacidade_cilindros()

index=0
ov_alocada=False
count_impossivel=0

print_acabamentos()
print_wip_inicial(id_ovs)



if method==1:

    for index in range(len(id_ovs)):
        id_ov=id_ovs[index]
        lista_items = ovs[id_ov].id_items
        for pos_item in range(len(lista_items)):

            id_item = lista_items[pos_item]

            lista_ofs = lista_ordenada(id_item)
            items[id_item].id_ofs = lista_ofs

            for pos_of in range(len(lista_ofs)):

                id_of = lista_ofs[pos_of]

                for pos_semana in range(len(ofs[id_of].id_alocada)):

                    print(str(ofs[id_of].cod_of) + ' ' + str(ofs[id_of].id_alocada) )
                    semana_desalocar=ofs[id_of].id_alocada[pos_semana]
                    duracao_desalocar=ofs[id_of].alocada_duracao[pos_semana]
                    desalocar(id_of,semana_desalocar,id_ov,duracao_desalocar)
                    ofs[id_of].id_alocada.remove(semana_desalocar)
                    ofs[id_of].alocada_duracao.remove(duracao_desalocar)

print_capacidade_reservada(0)

index=0

while index<(len(id_ovs)):

    id_ov=id_ovs[index]
    lista_items=ovs[id_ov].id_items
    semana_pretendida=ovs[id_ov].data_desejada #TODO VERIFICAR DATA MIN COM LEADTIME
    id_cliente=ovs[id_ov].id_cliente
    carga_completa=clientes[id_cliente].carga_completa
    items_completos=0
    remover_ov=0

    if semana_pretendida<0:
        semana_pretendida=0

    for pos_item in range(len(lista_items)):

        id_item=lista_items[pos_item]

        verificar_planeador(id_item)

        lista_ofs=lista_ordenada(id_item)
        items[id_item].id_ofs=lista_ofs

        for pos_of in range(len(lista_ofs)):

            semana_pretendida=ovs[id_ov].data_desejada
            if semana_pretendida < 0:
                semana_pretendida = 0

            id_of=lista_ofs[pos_of]

            semana_min_precedencia = semana_pretendida
            semana_max_precedencia=semana_pretendida

            for pos_precedente in range(len(ofs[id_of].precedencias)):
                id_precedencia = ofs[id_of].precedencias[pos_precedente]
                if len(ofs[id_precedencia].id_alocada)>0:

                    if min(ofs[id_precedencia].id_alocada) > semana_min_precedencia:
                        semana_min_precedencia = min(ofs[id_precedencia].id_alocada)

                    if max(ofs[id_precedencia].id_alocada)>semana_max_precedencia:
                        semana_max_precedencia=max(ofs[id_precedencia].id_alocada)

                else:
                    count_impossivel+=1

            semana_pretendida=semana_max_precedencia-ofs[id_of].n_semanas+1

            if semana_pretendida < 0:
                semana_pretendida=0

            if ofs[id_of].duracao==0:
                alocar(id_of,0,id_ov,ofs[id_of].duracao)
                ofs[id_of].semana_min = 0

            elif verificar_capacidade_rolante(id_of,semana_pretendida,id_ov,0.5,0.95)==True:

                alocar_capacidade_rolante(id_of, semana_pretendida, id_ov,0.5,0.95)
                ofs[id_of].semana_min=max(ofs[id_of].id_alocada)

            elif verificar_possivel_atras(id_of,semana_pretendida,id_ov,semana_min_precedencia)!=-1:

                semana_a_alocar=verificar_possivel_atras(id_of,semana_pretendida,id_ov,semana_min_precedencia)
                alocar_capacidade_rolante(id_of, semana_a_alocar, id_ov,0.5,0.95)
                ofs[id_of].semana_min = semana_a_alocar

            else:

                semana_impossivel=True

                id_ct=ofs[id_of].id_ct

                n_semanas=len(cts[id_ct].capacidade)-2

                while semana_impossivel == True and semana_pretendida<n_semanas:

                    semana_pretendida+=1

                    if verificar_capacidade_rolante(id_of,semana_pretendida,id_ov,0.95,ofs[id_of].duracao)==True:

                        semana_impossivel = False
                        alocar_capacidade_rolante(id_of, semana_pretendida, id_ov,0.5,0.95)
                        ofs[id_of].semana_min = semana_pretendida

                if semana_impossivel==True:
                    remover_ov=1
                    index+=1


        if carga_completa==0 and remover_ov==0:

            nova_semana=minimizar_wip_item(id_item,-1,parametro_duracao)

            items_completos += 1

        if items_completos==len(lista_items):
            index+=1

    if carga_completa==1 and remover_ov==0:

        nova_semana=minimizar_wip(id_ov,parametro_duracao)

        index += 1


print_output(method)
print_mapa_precedencias()
print_data_entrega(id_ovs)

n_ovs_possiveis=0

for index in range(len(id_ovs)):

    id_ov=id_ovs[index]

    n_ofs=0

    for pos_item in range(len(ovs[id_ov].id_items)):

        id_item=ovs[id_ov].id_items[pos_item]
        n_ofs+=len(items[id_item].id_ofs)

    if n_ofs>0:
        n_ovs_possiveis+=1

print('número de ovs ' + str(n_ovs_possiveis))


gerar_output_sap(id_ovs)
end = time.time()
print('Tempo de corrida: ' + str(((end - start)*60)) + ' segundos')









