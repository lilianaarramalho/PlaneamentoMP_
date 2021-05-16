class cliente(object):

    def __init__(self,id,nome,lead_time,carga_completa):

        self.id=id
        self.nome=nome
        self.lead_time=lead_time
        self.carga_completa=carga_completa
        self.descricao_clientes=[]

    def __repr__(self):
        return str(self.nome)

