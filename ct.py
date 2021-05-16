class ct(object):

    def __init__(self,id,nome,acabamento):

        self.id=id
        self.nome=nome
        self.acabamento=acabamento
        self.capacidade=[]
        self.capacidade_clientes=[[]]
        self.capacidade_blocos=[]
        self.capacidade_iniciais=[]
        self.capacidade_virados=[]

    def __repr__(self):
        return str(self.nome)
