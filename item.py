class item(object):

    def __init__(self,id,cod_item,id_ov,planeador):
        self.id=id
        self.id_ov=id_ov
        self.cod_item=cod_item
        self.id_ofs=[]
        self.planeador=planeador

    def __repr__(self):
        return str(self.id)