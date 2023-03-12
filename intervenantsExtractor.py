from pathlib import Path
import openpyxl
import pandas as pd
import os

def extractIntervenantsInformation(intervenantLine):
    identifiant = "".join([word[0].upper() for word in intervenantLine[4].value.split(" ")]) if intervenantLine[4].value else ""
    nom = " ".join([word.upper() for word in intervenantLine[4].value.split(" ") if word.isupper()]) if intervenantLine[4].value else ""
    prenom = " ".join([word.capitalize() for word in intervenantLine[4].value.split(" ") if not word.isupper()]) if intervenantLine[4].value else ""
    nom_uppercase = " ".join([word for word in nom.split() if word.isupper()])
    adresse = f"{prenom.lower().split()[0]}.{nom_uppercase.lower().replace(' ', '-')}@univ-rennes1.fr" if nom_uppercase and prenom else ""
    statut = ""
    employeur = ""
    
    return [identifiant, nom, prenom, adresse, statut, employeur]

def batchIntervenantsInformation(intervenantSheet, intervall=(2, None)):
    processedData = []
    for row in intervenantSheet.iter_rows(min_row=intervall[0], max_row=intervall[1]):
        processedData.append(extractIntervenantsInformation(row))

    return processedData

def DataFramesToExcel(
    dataframes: [pd.DataFrame], sheetNames: [str], output: str or Path
):
    if not os.path.exists(output):
        # Create the output directory if it doesn't exist
        os.makedirs(os.path.dirname(output), exist_ok=True)
    # Create an Excel file in which each sheet is associated with a dataframe in the list "dataframes" and its name is indicated in the list "sheetNames"

    with pd.ExcelWriter(output) as writer:
        for index in range(len(sheetNames)):
            dataframes[index].to_excel(
                writer, sheet_name=sheetNames[index], index=False
            )


### Main
def main(filePath):
    wb = openpyxl.load_workbook(filePath)
    intervenantSheet = wb.active
    intervenantsDF = batchIntervenantsInformation(intervenantSheet)

    # Remove duplicates based on nom and prenom columns
    intervenantsDF = pd.DataFrame(intervenantsDF, columns=["Identifiant", "Nom", "Prenom", "Adresse", "Statut", "Employeur"])
    intervenantsDF = intervenantsDF.drop_duplicates(subset=["Nom", "Prenom"], keep="first")
    
    DataFramesToExcel([intervenantsDF], ["Intervenants"], Path("output/intervenants.xlsx"))

if __name__ == "__main__":
    main("data/liste_intervenants.xlsx")