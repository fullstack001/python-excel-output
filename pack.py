import pandas as pd
import numpy as np
jaar = 2022
pad = 'D:/Pack/'
facturen_bron = pd.read_excel(pad + 'Export Factuurgegevens Qlik 2022.xlsx', sheet_name = 'Export facturen - Bewerkt V&V')
prijslijst_bron1 = pd.read_excel(pad + 'Prijs per product en debiteur - Export tabel VA-219 + VA-217 + AF-220.xlsx' , sheet_name = 'Debiteur-prijslijst')
prijslijst_bron2 = pd.read_excel(pad + 'Prijs per product en debiteur - Export tabel VA-219 + VA-217 + AF-220.xlsx' , sheet_name = 'va-217 Prijslijst-productprijs')
prijslijst_bron3 = pd.read_excel(pad + 'Prijs per product en debiteur - Export tabel VA-219 + VA-217 + AF-220.xlsx' , sheet_name = 'va-219 Debiteur-productprijs')

facturen = facturen_bron.copy()
facturen = facturen.loc[~facturen['Artikelnummer'].str.startswith('B')]
facturen['Artikelnummer'] = facturen['Artikelnummer'].astype(int)

prijslijst1 = pd.merge(prijslijst_bron2, prijslijst_bron1, how = 'left', left_on = 'cdprl', right_on = 'Prijslijst')
prijslijst1 = prijslijst1.rename(columns={"Debnr":"cddeb"})

prijslijst = pd.concat([prijslijst1,prijslijst_bron3])


prijslijst = prijslijst.loc[(prijslijst['datum-in'].dt.year == jaar),:]
prijslijst['Prijs per stuk'] = prijslijst['prijs']/100

for i in facturen.index:
    artnr = facturen.loc[i,'Artikelnummer']
    debnr = facturen.loc[i, 'Debnr']
    facdatum = facturen.loc[i, 'Invoice date']
    eenheid = facturen.loc[i, 'Unit of measure']
    try:
        facturen.loc[i, 'Prijsafspraak'] = prijslijst.loc[(prijslijst['cdstandeenhd'] == eenheid) & (prijslijst['cddeb'] == debnr) & (prijslijst['cdprodukt'] == artnr) & (prijslijst['datum-in'] <= facdatum) & (prijslijst['datum-uit'] >= facdatum),'Prijs per stuk'].values[0]
        #facturen.loc[i, 'Eenheid'] = prijslijst.loc[(prijslijst['cddeb'] == debnr) & (prijslijst['cdprodukt'] == artnr) & (prijslijst['datum-in'] <= facdatum) & (prijslijst['datum-uit'] >= facdatum),'cdstandeenhd'].values[0]
    except:
        pass

with pd.ExcelWriter(f'Contractanalyse Pack-it {jaar}.xlsx') as writer:
    facturen.to_excel(writer, index=False, sheet_name='Contractanalyse')
    prijslijst.to_excel(writer, index=False, sheet_name = 'Prijslijst')