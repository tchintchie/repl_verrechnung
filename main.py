#import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
#from matplotlib.patches import Circle, Rectangle, Arc
import seaborn as sns
#import pip
import subprocess as sp
import sys

#print(pip.__version__)
sp.check_call([sys.executable, '-m', 'pip', 'install', 'xlrd'])


vor_infra = pd.ExcelFile("vormonat/KO_Verrechnung_CeBrACloud_Infrastructure_OK.xlsx")
vor_extern = pd.ExcelFile("vormonat/KO_Verrechnung_CeBrACloudServer_extern_OK.xlsx")

neu_infra = pd.ExcelFile("neu/KO_Verrechnung_CeBrACloud_Infrastructure_OK.xlsx")
neu_extern = pd.ExcelFile("neu/KO_Verrechnung_CeBrACloudServer_extern_OK.xlsx")

ws_infra = vor_infra.sheet_names
ws_extern = vor_extern.sheet_names

print("Worksheets of Infrastructure Report: ", ws_infra)
print("Worksheets of Extern Report: ", ws_extern)


def make_df_from_sheets(xls):

    complete = pd.DataFrame()
    each_ws = pd.read_excel(xls, sheet_name=None)
    for name, data in each_ws.items():
      data["Profil"] = name
      complete = complete.append(data)

    return complete


vor_infra_complete = make_df_from_sheets(vor_infra)
vor_extern_complete = make_df_from_sheets(vor_extern)
neu_infra_complete = make_df_from_sheets(neu_infra)
neu_extern_complete = make_df_from_sheets(neu_extern)
print(neu_extern_complete.MID.head())
print(neu_infra_complete.MID.head())
print(vor_infra_complete.MID.head())
print(vor_extern_complete.MID.head())

# Gesamtsumme je Preisprofil je Report und Monat:
vor_profilsumme_infra = vor_infra_complete.groupby(["Profil"]).sum()
neu_profilsumme_infra = neu_infra_complete.groupby(["Profil"]).sum()
vor_profilsumme_extern = vor_extern_complete.groupby(["Profil"]).sum()
neu_profilsumme_extern = neu_extern_complete.groupby(["Profil"]).sum()


print("Gesamtsumme je Preisprofil Infrastracture je Monat:")
print(vor_profilsumme_infra)
print(neu_profilsumme_infra)


print("Gesamtsumme je Preisprofil Extern je Monat:")
print(vor_profilsumme_extern)
print(neu_profilsumme_extern)

# Kontrolle ob Summen übereinstimmen:
print("Kontrolle ob Summen überein stimmen:")

print("Profilsummen Infrastructure Vormonat: ", np.sum(vor_profilsumme_infra))
print("Profilsummen Infrastructure aktuell: ", np.sum(neu_profilsumme_infra))
print("Profilsummen Extern Vormonat: ", np.sum(vor_profilsumme_extern))
print("Profilsummen Extern aktuell: ", np.sum(neu_profilsumme_extern))
print()
print()

print("************************************************************************")
print("Gesamtsumme je Monat je Report: ")

vor_sum_infra = vor_infra_complete["Preis 2 (€)"].sum()
neu_sum_infra = neu_infra_complete["Preis 2 (€)"].sum()
print("************************************************************************")
print()
print()
vor_sum_extern = vor_extern_complete["Preis 2 (€)"].sum()
neu_sum_extern = neu_extern_complete["Preis 2 (€)"].sum()

print("Summe Infrastructure Vormonat: ", vor_sum_infra)
print("Summe Infrastructure aktuell: ", neu_sum_infra)
print("Summe Extern Vormonat: ", vor_sum_extern)
print("Summe Extern aktuell: ", neu_sum_extern)
print()
print()
print("************************************************************************")
print("Differenzen je Preisprofil zum Vormonat:")
print("************************************************************************")
print()
print()
print("Differnz Infrastructure gesamt zum Vormonat: ",round(neu_sum_infra-vor_sum_infra))
print("************************************************************************")
print("Differenz je Profil Infrastructure zum Vormonat: ", round(neu_profilsumme_infra - vor_profilsumme_infra))
print("************************************************************************")
print()
print()
print("Differenz Extern gesamt zum Vormonat: ", round(neu_sum_extern - vor_sum_extern))
print("************************************************************************")
print("Differenz je Profil Extern zum Vormonat: ", round(neu_profilsumme_extern - vor_profilsumme_extern))

print("************************************************************************")
# Graphing the differences of infrastructure reports:
# merge the two datasets and assign column to distinguish them

infra_vergleich = pd.concat([vor_profilsumme_infra.assign(Monat='Vormonat'), neu_profilsumme_infra.assign(Monat='aktueller Monat')])

infra_vergleich.reset_index(inplace=True)

fig, ax = plt.subplots(figsize = (15,15))
ax = sns.catplot(x="Profil", y="Preis 2 (€)", hue="Monat", data=infra_vergleich, height=8, kind="bar", palette="muted", legend=False)


ax.set_title='Infrastructure Vergleich zu Vormonat'
ax.set_ylabels("Umsatz in EURO")
ax.set_xlabels("Profile")

plt.legend(loc="upper left")
fig.tight_layout()

plt.savefig("diff_infra.png")
plt.clf()


# Graphing the differences of extern reports:
# merge the two datasets and assign column to distinguish them

extern_vergleich = pd.concat([vor_profilsumme_extern.assign(Monat='Vormonat'), neu_profilsumme_extern.assign(Monat='aktueller Monat')])

extern_vergleich.reset_index(inplace=True)

fig, ax = plt.subplots(figsize = (15,15))
ax = sns.catplot(x="Profil", y="Preis 2 (€)", hue="Monat", data=extern_vergleich, height=8, kind="bar", palette="muted", legend=False)

ax.set_title='Extern Vergleich zu Vormonat'
ax.set_ylabels("Umsatz in EURO")
ax.set_xlabels("Profile")

plt.legend(loc="upper left")
fig.tight_layout()

plt.savefig("diff_extern.png")
plt.clf()


