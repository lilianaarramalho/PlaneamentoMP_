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
import warnings
warnings.filterwarnings("ignore")
from operator import itemgetter, attrgetter, methodcaller

lista_verdes,lista_azuis=ler_tintas_alocadas()

alterar_centro_trabalho_amoltex()

criar_capacidade_rolante()

verificar_referencias_c()

ler_cts_embalagem()

import_bom()

calcular_n_ofs(0.5)

parametro_duracao=60

verificar_precedencias()

temp_ovs=sorted(ovs, key=attrgetter('data_desejada', 'data_criacao'),reverse=True).copy()

id_ovs=get_ids(temp_ovs)

print_capacidade_reservada(method)

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

id_items_to_remove=[]
id_ofs_to_remove=[]
id_ovs_to_remove=[]

while index<(len(id_ovs)):

    id_ov=id_ovs[index]
    lista_items=ovs[id_ov].id_items
    semana_pretendida=ovs[id_ov].data_desejada #TODO VERIFICAR DATA MIN COM LEADTIME
    id_cliente=ovs[id_ov].id_cliente
    carga_completa=clientes[id_cliente].carga_completa
    items_completos=0
    remover_ov=0

    print('A alocar OV: ' + str(ovs[id_ov].cod_ov)  + ' com semana pretendida ' + str(semana_pretendida+semana_inicio_plano))

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

            elif verificar_capacidade_rolante(id_of,semana_pretendida,id_ov,0.5,0.95,lista_verdes,lista_azuis)==True:

                alocar_capacidade_rolante(id_of, semana_pretendida, id_ov,0.5,0.95,lista_verdes,lista_azuis)
                ofs[id_of].semana_min=max(ofs[id_of].id_alocada)

            elif verificar_possivel_atras(id_of,semana_pretendida,id_ov,semana_min_precedencia,lista_verdes,lista_azuis)!=-1:

                print('A verificar se ?? poss??vel produzir ' + str(ovs[id_ov].cod_ov) + ' na semana anterior ' + str(semana_pretendida+semana_inicio_plano))

                semana_a_alocar=verificar_possivel_atras(id_of,semana_pretendida,id_ov,semana_min_precedencia, lista_verdes, lista_azuis)
                alocar_capacidade_rolante(id_of, semana_a_alocar, id_ov,0.5,0.95,lista_verdes,lista_azuis)
                ofs[id_of].semana_min = semana_a_alocar

            else:

                semana_impossivel=True

                id_ct=ofs[id_of].id_ct

                n_semanas=len(cts[id_ct].capacidade)-2

                print('A verificar se ?? poss??vel produzir ' + str(ovs[id_ov].cod_ov) + ' na semana seguinte ' + str(semana_pretendida+1+semana_inicio_plano))

                while semana_impossivel == True and semana_pretendida<n_semanas:

                    semana_pretendida+=1

                    if verificar_capacidade_rolante(id_of,semana_pretendida,id_ov,0.95,ofs[id_of].duracao,lista_verdes,lista_azuis)==True:

                        semana_impossivel = False
                        alocar_capacidade_rolante(id_of, semana_pretendida, id_ov,0.5,0.95,lista_verdes,lista_azuis)
                        ofs[id_of].semana_min = semana_pretendida

                if semana_impossivel==True:
                    remover_ov=1
                    index+=1

                    id_items = ovs[id_ov].id_items

                    for id_item in id_items:

                        id_ofs = items[id_item].id_ofs

                        for id_of in id_ofs:
                            id_ofs_to_remove.append(id_of)

                        id_items_to_remove.append(id_item)

                    id_ovs_to_remove.append(id_ov)

                    break


        if carga_completa==0 and remover_ov==0:

            nova_semana=minimizar_wip_item(id_item,-1,parametro_duracao)

            print('A minimizar wip do Item')

            print('--------------------------------------------------')

            items_completos += 1

        if items_completos==len(lista_items):
            index+=1

    if carga_completa==1 and remover_ov==0:

        print('A minimizar wip da OV')

        print('--------------------------------------------------')

        nova_semana=minimizar_wip(id_ov,parametro_duracao)

        index += 1

# count=0
# total=len(items)
# while count<total:
#     if items[count].id in id_items_to_remove:
#         items.remove(items[count])
#         total=len(items)
#     else:
#         count+=1
#
# count=0
# total=len(id_items)
# while count<total:
#     if id_items[count] in id_items_to_remove:
#         id_items.remove(id_items[count])
#         total = len(id_items)
#     else:
#         count+=1
#
# count = 0
# total = len(ofs)
# while count < total:
#     if ofs[count].id in id_ofs_to_remove:
#         ofs.remove(ofs[count])
#         total = len(ofs)
#     else:
#         count += 1
#
# count = 0
# total = len(id_ofs)
# while count < total:
#     if id_ofs[count] in id_ofs_to_remove:
#         id_ofs.remove(id_ofs[count])
#         total=len(id_ofs)
#     else:
#         count += 1
#
# count = 0
# total = len(ovs)
# while count < total:
#     if ovs[count].id in id_ovs_to_remove:
#         ovs.remove(ovs[count])
#         total = len(ovs)
#     else:
#         count += 1
#
# count = 0
# total = len(id_ovs)
# while count < total:
#     if id_ovs[count] in id_ovs_to_remove:
#         id_ovs.remove(id_ovs[count])
#         total = len(id_ovs)
#     else:
#         count += 1

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

print('n??mero de ovs planeadas: ' + str(n_ovs_possiveis))
print('--------------------------------------------------')


gerar_output_sap(id_ovs)
end = time.time()

print('Tempo de corrida: ' + str((round((end - start))*60)) + ' segundos')
print('--------------------------------------------------')

