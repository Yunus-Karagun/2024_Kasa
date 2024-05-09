import pandas as pd
import numpy as np
import os
import warnings
warnings.filterwarnings('ignore')

Parametre = pd.read_excel("D:\\Kasalar\\Config\\Config.xlsx",  sheet_name= "Parametre", index_col=None) #Kasa Föyü Adresleri
magaza_kasa = pd.read_excel("D:\\Kasalar\\Config\\Config.xlsx",  sheet_name= "Adres", index_col=0) #Kasa Föyü Adresleri
kodlar = pd.read_excel("D:\\Kasalar\\Config\\Config.xlsx",  sheet_name= "KF",  index_col=None) #Kasa Kodları
hucre = magaza_kasa["Adres"]

#Parametreler
gun=Parametre.iloc[0][1] #Kasa Dosyası yapılacak aydaki gün sayısı yazılacak
ay=Parametre.iloc[1][1] #hangi ay olduğu
yil=Parametre.iloc[2][1]
Konum=Parametre.iloc[3][1]
Rapor=Parametre.iloc[4][1]
loc = (Konum + ay + "\\") #dosyaların bulunduğu adres
N_Gunu = str(gun).zfill(2)+"."+ ay + "."+ yil #"08.09.2021" #Raporda


sayfalar = []  #excel dosyasındaki tarih olan sayfa adlarından oluşan bir liste yapıyoruz.
for gn in range (1, gun+1):
    sayfalar.append((str(gn).zfill(2)+"."+ ay + "."+ yil))

for filename in os.listdir(loc):
    ciro = pd.read_excel(os.path.join(directory, filename),
                      sheet_name=sayfalar,
                      skiprows = 2, nrows = 65, header = None, index_col=None, usecols=[0, 1, 2, 3, 4, 5, 6, 7])

    for g in range (0, gun):
        gunluk_kasa = []
        gunluk_kasa.append(list(ciro.keys())[g])
        for i in range (1, 67):
            aa=ciro[list(ciro.keys())[g]].iloc[eval(hucre[i])]
            gunluk_kasa.append(aa)  
        magaza_kasa[filename + "-" + str(g+1).zfill(2)] = gunluk_kasa     

magaza_kasa_T = magaza_kasa.T
magaza_kasa_T['MAĞAZA ADI'] = magaza_kasa_T['MAĞAZA ADI'].str.strip()
magaza_kasa_T = magaza_kasa_T.merge(kodlar,
                                      on ='MAĞAZA ADI', how="left").set_axis(magaza_kasa_T.index)

# Yatırlan Nakitler Bölümü

df2 = pd.DataFrame()
for filename in os.listdir(loc):
    sheets = pd.read_excel(os.path.join(directory, filename),
                      sheet_name=sayfalar,
                      skiprows = 57, nrows = 5, header = 1, index_col=0, usecols=[0, 1, 2, 3, 4, 5, 6])
    
    magaza = pd.read_excel(os.path.join(directory, filename),
                           sheet_name=sayfalar[0],
                           skiprows = 2, nrows = 1, header = None, usecols=[5])
    
    dfs = []
    for framename in sheets.keys():
        temp_df = sheets[framename]
        temp_df['Sayfa'] = framename
        dfs.append(temp_df)
    df = pd.concat(dfs)
    df.loc[:,"Dosya Adı"] = str(i).zfill(3) + ".xlsx"
    df.loc[:,"Magaza"] = magaza.at[0, 5]
    
 
    df1 = df.dropna(thresh=4)
    df2 = pd.concat([df2, df1])

df2['Magaza'] = df2['Magaza'].str.strip()
g_fark = magaza_kasa_T[["MAĞAZA ADI", "DS-Tarih", "GENEL FARK"]]
g_fark.drop("Adres", inplace=True)
g_fark[(g_fark["GENEL FARK"] < -1) | (g_fark["GENEL FARK"] > 1)]
Nakitler = magaza_kasa_T[["KASA KODU", "MAĞAZA ADI", "DS-Tarih", "FATURALI SATIŞLAR", "GİDER PUSULASI", "CİRO TOPLAMI", "NAKİT_1", "EURO TUTAR", "USD TUTAR", "MASRAF", "NAKİT", "YATIRILAN TL FARK", "YATIRILAN EURO FARK","YATIRILAN USD FARK", "İL"]]
Nakitler.drop("Adres", inplace=True)

Bugun_n = Nakitler[Nakitler["DS-Tarih"] == N_Gunu]
Bugun_n1 = Bugun_n[["DS-Tarih", "KASA KODU", "MAĞAZA ADI", "NAKİT_1", "EURO TUTAR", "USD TUTAR"]]
Nakitler1=Nakitler[["DS-Tarih", "KASA KODU", "MAĞAZA ADI", "NAKİT_1", "EURO TUTAR", "USD TUTAR"]]
with pd.ExcelWriter(Rapor+N_Gunu+'-x-Rapor.xlsx') as writer:
    magaza_kasa_T.to_excel(writer, sheet_name = 'Raw')   
    Nakitler.to_excel(writer, sheet_name = 'Main')
    Bugun_n.to_excel(writer, sheet_name = N_Gunu)
    Bugun_n1.to_excel(writer, sheet_name = N_Gunu+"-Nakit")
    Nakitler1.to_excel(writer, sheet_name = 'Main-Nakit')
    g_fark[(g_fark["GENEL FARK"] < -1) | (g_fark["GENEL FARK"] > 1)].to_excel(writer, sheet_name = "Hatalı")
    magaza_kasa_T[["MAĞAZA ADI", "DS-Tarih", "NOTLAR"]].dropna().to_excel(writer, sheet_name = "Notlar")
    df2[(df2[["YATIRILAN TL", "YATIRILAN USD", "YATIRILAN EURO"]].sum(axis=1, skipna=True) != 0)].to_excel(writer, sheet_name = "Yatırılan Nakitler")

    
g_fark = magaza_kasa_T[["MAĞAZA ADI", "DS-Tarih", "GENEL FARK"]]
g_fark.drop("Adres", inplace=True)
g_fark[(g_fark["GENEL FARK"] < -1) | (g_fark["GENEL FARK"] > 1)]
Nakitler = magaza_kasa_T[["KASA KODU", "MAĞAZA ADI", "DS-Tarih", "FATURALI SATIŞLAR", "GİDER PUSULASI", "CİRO TOPLAMI", "NAKİT_1", "EURO TUTAR", "USD TUTAR", "MASRAF", "NAKİT", "YATIRILAN TL FARK", "YATIRILAN EURO FARK","YATIRILAN USD FARK", "İL"]]
Nakitler.drop("Adres", inplace=True)

Bugun_n = Nakitler[Nakitler["DS-Tarih"] == N_Gunu]
Bugun_n1 = Bugun_n[["DS-Tarih", "KASA KODU", "MAĞAZA ADI", "NAKİT_1", "EURO TUTAR", "USD TUTAR"]]
Nakitler1=Nakitler[["DS-Tarih", "KASA KODU", "MAĞAZA ADI", "NAKİT_1", "EURO TUTAR", "USD TUTAR"]]
with pd.ExcelWriter(Rapor+"control.xlsx") as writer:
    magaza_kasa_T.to_excel(writer, sheet_name = 'Raw')   
    Nakitler.to_excel(writer, sheet_name = 'Main')
    Bugun_n.to_excel(writer, sheet_name = N_Gunu)
    Bugun_n1.to_excel(writer, sheet_name = N_Gunu+"-Nakit")
    Nakitler1.to_excel(writer, sheet_name = 'Main-Nakit')
    g_fark[(g_fark["GENEL FARK"] < -1) | (g_fark["GENEL FARK"] > 1)].to_excel(writer, sheet_name = "Hatalı")
    magaza_kasa_T[["MAĞAZA ADI", "DS-Tarih", "NOTLAR"]].dropna().to_excel(writer, sheet_name = "Notlar")
    df2[(df2[["YATIRILAN TL", "YATIRILAN USD", "YATIRILAN EURO"]].sum(axis=1, skipna=True) != 0)].to_excel(writer, sheet_name = "Yatırılan Nakitler")
    

