class of(object):

    def __init__(self,id,cod_of,duracao,quantidade,id_ct,codigo_material,descricao_material,codigo_precedencia,id_items,blocos,viradas):

        self.id=id
        self.cod_of=cod_of
        self.duracao=duracao
        self.quantidade=quantidade
        self.id_ct=id_ct
        self.codigo_material=codigo_material
        self.descricao_material=descricao_material
        self.codigo_precedencia=codigo_precedencia
        self.id_items=id_items
        self.precedencias=[]
        self.sequencias=[]
        self.id_alocada=[]
        self.blocos=blocos
        self.viradas=viradas

    def __repr__(self):
        return str(self.cod_of)
