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
vor_intern = pd.ExcelFile("vormonat/KO_Verrechnung_CeBrACloudServer_intern_OK.xlsx")

neu_infra = pd.ExcelFile("neu/KO_Verrechnung_CeBrACloud_Infrastructure_OK.xlsx")
neu_extern = pd.ExcelFile("neu/KO_Verrechnung_CeBrACloudServer_extern_OK.xlsx")
neu_intern = pd.ExcelFile("neu/KO_Verrechnung_CeBrACloudServer_intern_OK.xlsx")

ws_infra = vor_infra.sheet_names
ws_extern = vor_extern.sheet_names
ws_intern = vor_intern.sheet_names

print("Worksheets of Infrastructure Report: ", ws_infra)
print("Worksheets of Extern Report: ", ws_extern)
print("Worksheets of Intern Report: ", ws_intern)


def make_df_from_sheets(xls):

    complete = pd.DataFrame()
    each_ws = pd.read_excel(xls, sheet_name=None)
    for name, data in each_ws.items():
        data["Profil"] = name
        complete = complete.append(data)

    return complete


vor_infra_complete = make_df_from_sheets(vor_infra)
vor_extern_complete = make_df_from_sheets(vor_extern)
vor_intern_complete = make_df_from_sheets(vor_intern)
neu_infra_complete = make_df_from_sheets(neu_infra)
neu_extern_complete = make_df_from_sheets(neu_extern)
neu_intern_complete = make_df_from_sheets(neu_intern)
print(neu_extern_complete.MID.head())
print(neu_infra_complete.MID.head())
print(neu_intern_complete.MID.head())
print(vor_infra_complete.MID.head())
print(vor_extern_complete.MID.head())
print(vor_intern_complete.MID.head())

# Gesamtsumme je Preisprofil je Report und Monat:
vor_profilsumme_infra = vor_infra_complete.groupby(["Profil"]).sum()
neu_profilsumme_infra = neu_infra_complete.groupby(["Profil"]).sum()
vor_profilsumme_extern = vor_extern_complete.groupby(["Profil"]).sum()
neu_profilsumme_extern = neu_extern_complete.groupby(["Profil"]).sum()
vor_profilsumme_intern = vor_intern_complete.groupby(["Profil"]).sum()
neu_profilsumme_intern = neu_intern_complete.groupby(["Profil"]).sum()

print("Gesamtsumme je Preisprofil Infrastracture je Monat:")
print(vor_profilsumme_infra)
print(neu_profilsumme_infra)

print("Gesamtsumme je Preisprofil Extern je Monat:")
print(vor_profilsumme_extern)
print(neu_profilsumme_extern)

print("Gesamtsumme je Preisprofil Intern je Quartal:")
print(vor_profilsumme_intern)
print(neu_profilsumme_intern)



# Kontrolle ob Summen übereinstimmen:
print("Kontrolle ob Summen überein stimmen:")

print("Profilsummen Infrastructure Vormonat: ", np.sum(vor_profilsumme_infra))
print("Profilsummen Infrastructure aktuell: ", np.sum(neu_profilsumme_infra))
print("Profilsummen Extern Vormonat: ", np.sum(vor_profilsumme_extern))
print("Profilsummen Extern aktuell: ", np.sum(neu_profilsumme_extern))
print("Profilsummen Intern Vorquartal: ", np.sum(vor_profilsumme_intern))
print("Profilsummen Intern aktuell: ", np.sum(neu_profilsumme_intern))
print()
print()

print("************************************************************************")
print("Gesamtsumme je Monat je Report: ")

vor_sum_infra = vor_infra_complete["Preis 1 (€)"].sum()
neu_sum_infra = neu_infra_complete["Preis 1 (€)"].sum()
print("************************************************************************")
print()
print()
vor_sum_extern = vor_extern_complete["Preis 1 (€)"].sum()
neu_sum_extern = neu_extern_complete["Preis 1 (€)"].sum()

print("************************************************************************")
print()
print()
vor_sum_intern = vor_intern_complete["Preis 1 (€)"].sum()
neu_sum_intern = neu_intern_complete["Preis 1 (€)"].sum()

print("Summe Infrastructure Vormonat: ", vor_sum_infra)
print("Summe Infrastructure aktuell: ", neu_sum_infra)
print("Summe Extern Vormonat: ", vor_sum_extern)
print("Summe Extern aktuell: ", neu_sum_extern)
print("Summe Intern Vorquartal: ", vor_sum_intern)
print("Summe Intern aktuell: ", neu_sum_intern)

print()
print()
print("************************************************************************")
print("Differenzen je Preisprofil zum Vormonat:")
print("************************************************************************")
print()
print()
print("Differnz Infrastructure gesamt zum Vormonat: ",
      round(neu_sum_infra - vor_sum_infra))
print("************************************************************************")
print("Differenz je Profil Infrastructure zum Vormonat: ",
      round(neu_profilsumme_infra - vor_profilsumme_infra))
print("************************************************************************")
print()
print()
print("Differenz Extern gesamt zum Vormonat: ",
      round(neu_sum_extern - vor_sum_extern))
print("************************************************************************")
print("Differenz je Profil Extern zum Vormonat: ",
      round(neu_profilsumme_extern - vor_profilsumme_extern))

print("************************************************************************")
print()
print()


print("Differenz Intern gesamt zum Vorquartal: ",
      round(neu_sum_intern - vor_sum_intern))
print(
    "************************************************************************")
print("Differenz je Profil Intern zum Vorquartal: ",
      round(neu_profilsumme_intern - vor_profilsumme_intern))

print(
    "************************************************************************")


# Graphing the differences of infrastructure reports:
# merge the two datasets and assign column to distinguish them

infra_vergleich = pd.concat([
    vor_profilsumme_infra.assign(Monat='Vormonat'),
    neu_profilsumme_infra.assign(Monat='aktueller Monat')
])

infra_vergleich.reset_index(inplace=True)

fig, ax = plt.subplots(figsize=(15, 15))
ax = sns.catplot(
    x="Profil",
    y="Preis 1 (€)",
    hue="Monat",
    data=infra_vergleich,
    height=8,
    kind="bar",
    palette="muted",
    legend=False)

ax.set_title = 'Infrastructure Vergleich zu Vormonat'
ax.set_ylabels("Umsatz in EURO")
ax.set_xlabels("Profile")

plt.legend(loc="upper left")
fig.tight_layout()

plt.savefig("diff_infra.png")
plt.clf()

# Graphing the differences of extern reports:
# merge the two datasets and assign column to distinguish them

extern_vergleich = pd.concat([
    vor_profilsumme_extern.assign(Monat='Vormonat'),
    neu_profilsumme_extern.assign(Monat='aktueller Monat')
])

extern_vergleich.reset_index(inplace=True)

fig, ax = plt.subplots(figsize=(15, 15))
ax = sns.catplot(
    x="Profil",
    y="Preis 1 (€)",
    hue="Monat",
    data=extern_vergleich,
    height=8,
    kind="bar",
    palette="muted",
    legend=False)

ax.set_title = 'Extern Vergleich zu Vormonat'
ax.set_ylabels("Umsatz in EURO")
ax.set_xlabels("Profile")

plt.legend(loc="upper left")
fig.tight_layout()

plt.savefig("diff_extern.png")
plt.clf()


# Graphing the differences of intern reports:
# merge the two datasets and assign column to distinguish them

intern_vergleich = pd.concat([
    vor_profilsumme_intern.assign(Monat='Vorquartal'),
    neu_profilsumme_intern.assign(Monat='aktuelles Quartal')
])

intern_vergleich.reset_index(inplace=True)

fig, ax = plt.subplots(figsize=(15, 15))
ax = sns.catplot(
    x="Profil",
    y="Preis 1 (€)",
    hue="Monat",
    data=intern_vergleich,
    height=8,
    kind="bar",
    palette="muted",
    legend=False)

ax.set_title = 'Intern Vergleich zu Vorquartal'
ax.set_ylabels("Umsatz in EURO")
ax.set_xlabels("Profile")

plt.legend(loc="upper left")
fig.tight_layout()

plt.savefig("diff_intern.png")
plt.clf()

