class ov(object):

    def __init__(self,id,cod_ov,id_cliente,id_interno,data_desejada,data_criacao,sold_to):
        self.id=id
        self.cod_ov=cod_ov
        self.id_cliente=id_cliente
        self.id_interno=id_interno
        self.data_desejada=data_desejada
        self.id_items=[]
        self.data_criacao=data_criacao
        self.data_min=data_desejada
        self.sold_to=sold_to

    def __repr__(self):
        return str(self.cod_ov) + ' ' + str(self.data_criacao)

