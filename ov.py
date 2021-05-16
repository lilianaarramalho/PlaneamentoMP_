class ov(object):

    def __init__(self,id,cod_ov,id_cliente,id_interno,data_desejada):
        self.id=id
        self.cod_ov=cod_ov
        self.id_cliente=id_cliente
        self.id_interno=id_interno
        self.data_desejada=data_desejada
        self.id_items=[]

    def __repr__(self):
        return str(self.cod_ov) + ' ' + str(self.semana)

