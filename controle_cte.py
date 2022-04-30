import pandas as pd
from xml.dom import minidom
from tkinter.filedialog import askopenfilenames
from datetime import datetime

EmD = []
NumNF = []
NumCTE = []
ClientName = []
DeliveryCity = []
CteValue = []

cte_path = askopenfilenames(filetypes=[("Text files", ".xml")])

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

for pos, each_date in enumerate(EmD):
    emission_date = datetime.fromisoformat(each_date)
    formatted_emission_date = f'{emission_date.day}/{emission_date.month:02d}/{emission_date.year}'
    EmD[pos] = formatted_emission_date

for pos, each_numNf in enumerate(NumNF):
    each_numNf = each_numNf[28:34]
    NumNF[pos] = each_numNf

cte_df = pd.DataFrame(list(zip(EmD, NumNF, NumCTE, ClientName, DeliveryCity, CteValue)), columns=['Data Emissao', 'Numero NF', 'Numero CTE', 'Nome Cliente', 'Cidade Entrega', 'Valor CTE'])

CteValueSum = list(map(float, CteValue))
total_cte = sum(CteValueSum)
cte_df.at[0, 'Valor Total das CTEs'] = total_cte
cte_df.index += 1

cte_df.to_excel("controle_cte.xlsx")