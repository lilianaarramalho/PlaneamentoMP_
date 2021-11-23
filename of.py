class of(object):

    def __init__(self,id,cod_of,duracao,quantidade,id_ct,codigo_material,descricao_material,codigo_precedencia,id_items,blocos,viradas,descricao_precedencia, quantidade_precedencia,tipo_ordem,planeador,descritivo_ct):

        self.id=id
        self.cod_of=cod_of
        self.duracao=float(duracao)
        self.quantidade=float(quantidade)
        self.id_ct=id_ct
        self.codigo_material=codigo_material
        self.descricao_material=descricao_material
        self.codigo_precedencia=[]
        self.codigo_precedencia.append(codigo_precedencia)
        self.id_items=id_items
        self.precedencias=[]
        self.sequencias=[]
        self.id_alocada=[]
        self.id_alocada_anterior=[]
        self.alocada_duracao_anterior=[]
        self.alocada_duracao=[]
        self.blocos=float(blocos)
        self.viradas=float(viradas)
        self.semana_min=-1
        self.semana_max=-1
        self.n_semanas=-1
        self.descricao_precedencia=descricao_precedencia
        self.quantidade_precedencia=quantidade_precedencia
        self.tipo_ordem=tipo_ordem
        self.planeador=planeador
        self.descritivo_ct=descritivo_ct
        self.cod_material=descricao_material.split(' ')[2].split('/')[0]
        try:
            self.acabamento = descricao_material.split('/')[1].split(' ')[0]
        except:
            self.acabamento = ""
        self.dim1=descricao_material.split(' ')[3].split('X')[0]
        self.plyup=0
        self.calandrado=0
        self.is_c=0
        self.tinta=-1

        if "X" in descricao_material:
            self.dim2 = descricao_material.split('X')[1].split(' ')[0]
        else:
            self.dim2=descricao_material

    def __repr__(self):
        return str(self.cod_of)

