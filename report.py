#Importanto Bibliotecas | Importing Library

from docxtpl import DocxTemplate, InlineImage
import pandas as pd
from random import randint
from docx.shared import Cm, Inches, Mm, Emu
from datetime import date
from docx2pdf import convert

#Lendo o template e os dados da Planilha (caso haja) | Reading the template and data from the Spreadsheet (if any)

doc = DocxTemplate("report_template.docx")

#Criando um DataSet Hipotético | Creating a Hypothetical DataSet

salesRows = []
list_item = ['Chairs', 'Storage', 'Phones', 'Tables', 'Accessories']
for iItr in range(5):
    costPu = randint(1,15)
    nUnits = randint(100,500)
    salesRows.append({'sNo': iItr+1, 'name': list_item[iItr],'nUnits': nUnits, 'cPU': costPu,  'revenue': costPu*nUnits})

print(salesRows)

topSalesItem = pd.DataFrame.from_dict(salesRows).nlargest(n= 3, columns="revenue").to_dict('records')
print(topSalesItem)

#Criando os Gráficos | Creating Graphics

import matplotlib.pyplot as plt
import numpy as np

revenue_2020 = list()
revenue_2021 = list()
revenue_2022 = list()

for x in range(1, 5):
    revenue_2020.append(randint(1000, 10000))
    revenue_2021.append(randint(1000, 10000))
    revenue_2022.append(randint(1000, 10000))

barWidth = 0.20

plt.figure(figsize=(5,5))

r1 = np.arange(len(revenue_2020))
r2 = [x + barWidth for x in r1]
r3 = [x + barWidth for x in r2]

plt.bar(r1, revenue_2020, color='#94CD29', width=barWidth, label='2020')
plt.bar(r2, revenue_2021, color='#273152', width=barWidth, label='2021')
plt.bar(r3, revenue_2022, color='#00BFFF', width=barWidth, label='2022')

#plt.style.use('bmh')
plt.xticks([r + barWidth for r in range(len(revenue_2020))], ['1º Quarter', '2º Quarter', '3º Quarter', '4º Quarter'])
plt.ylabel('Revenue')
plt.title('Result of the last 3 years', x=0.5, y=1.1)

plt.legend(bbox_to_anchor=(1.02, 1), borderaxespad=0)
#plt.show()
plt.savefig('result_last_3years.png', bbox_inches='tight')

revenue_acum = [sum(revenue_2020), sum(revenue_2021), sum(revenue_2022)]
print(revenue_acum)

year = date.today().year - 3 + revenue_acum.index(max(revenue_acum)) + 1 
print(year)

#Transferindo para o Word e Salvando | Transferring to Word and Saving

context = {
    "tblSalesRows": salesRows,
    "texSalesTotal": sum(revenue_2022),
    "numTopSalesItem": topSalesItem,
    "betterYear": year,
    "resulBetterYear": round(sum(revenue_2022)/sum(revenue_acum)*100, 2),
    "grafico1": InlineImage(doc, 'result_last_3years.png')
}

doc.render(context)
doc.save('Annual Report - Genereted.docx')
convert("Annual Report - Genereted.docx") #Convertendo para PDF | Converting to PDF

print('Done!')