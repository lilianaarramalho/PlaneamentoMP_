
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

verificar_precedencias()

id_ovs=sort_by_leadtime()

print_capacidade_reservada()

for index in range(len(id_ovs)):

    id_ov=id_ovs[index]
    lista_items=ovs[id_ov].id_items
    semana_pretendida=ovs[id_ov].data_min
    id_cliente=ovs[id_ov].id_cliente
    carga_completa=clientes[id_cliente].carga_completa


    if semana_pretendida<0:
        semana_pretendida=0

    for pos_item in range(len(lista_items)):

        id_item=lista_items[pos_item]

        lista_ofs=items[id_item].id_ofs

        for pos_of in range(len(lista_ofs)-1,0-1,-1):

            id_of=lista_ofs[pos_of]

            if len(ofs[id_of].precedencias)>0:
                id_precedencia=ofs[id_of].precedencias[0]

            if ofs[id_of].id_alocada==-1:

                if verificar_capacidades(id_of,semana_pretendida,id_ov,0.95)==True:

                    alocar(id_of, semana_pretendida, id_ov)

                elif verificar_possivel_atras(id_of,semana_pretendida,id_ov)!=-1:

                    semana_a_alocar=verificar_possivel_atras(id_of,semana_pretendida,id_ov)
                    alocar(id_of, semana_a_alocar, id_ov)

                else:

                    semana_impossivel=True

                    id_ct=ofs[id_of].id_ct

                    n_semanas=len(cts[id_ct].capacidade)-semana_pretendida-1

                    while semana_impossivel == True and semana_pretendida<n_semanas:
                        semana_pretendida+=1
                        if verificar_capacidades(id_of,semana_pretendida,id_ov,0.95)==True:
                            semana_impossivel = False
                            alocar(id_of, semana_pretendida, id_ov)

        if carga_completa==0:

            minimizar_wip_item(id_item,-1)

    if carga_completa==1:

        minimizar_wip(id_ov)

print_output()
print_mapa_precedencias()











