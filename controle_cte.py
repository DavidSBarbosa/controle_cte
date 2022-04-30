import pandas as pd
from xml.dom import minidom
from tkinter.filedialog import askopenfilenames

EmD = []
NumNF = []
NumCTE = []
ClientName = []
DeliveryCity = []
CteValue = []

cte_path = askopenfilenames()

for actual_xml in cte_path:
    tree = minidom.parse(actual_xml)

    emission_date = tree.getElementsByTagName('dhEmi')
    num_nf = tree.getElementsByTagName('chave')
    num_cte = tree.getElementsByTagName('nCT')
    client_name = tree.getElementsByTagName('xNome')
    delivery_city = tree.getElementsByTagName('xMun')
    cte_value = tree.getElementsByTagName('vRec')

    EmD.append(emission_date[0].firstChild.data)
    NumNF.append(num_nf[0].firstChild.data)
    NumCTE.append(num_cte[0].firstChild.data)
    ClientName.append(client_name[2].firstChild.data)
    DeliveryCity.append(delivery_city[2].firstChild.data)
    CteValue.append(cte_value[0].firstChild.data)

cte_df = pd.DataFrame(list(zip(EmD, NumNF, NumCTE, ClientName, DeliveryCity, CteValue)), columns=['Data Emissao', 'Numero NF', 'Numero CTE', 'Nome Cliente', 'Cidade Entrega', 'Valor CTE'])
cte_df.to_excel("controle_cte.xlsx")