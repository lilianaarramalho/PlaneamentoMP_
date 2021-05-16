
from import_data import *
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

for index in range(len(ovs)):
    id_cliente=ovs[index].id_cliente
    if clientes[id_cliente].lead_time!=0:
        print('debug')