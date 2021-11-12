import os
import pandas as pd

from constants import *
from various_functions import *


def file_treatment_C13(path_folder, name_file):
    print("TRAITEMENT DU FICHIER C13: " + name_file)
    DAYS = "lundi,mardi,mercredi,jeudi,vendredi,samedi,dimanche"
    START_NUMBER_COLUMN = 21

    dtype = {12: str}
    for i in range(21, 49):
        dtype[i] = str
    for i in range(55, 57):
        dtype[i] = str
    for i in range(58, 63):
        dtype[i] = str
    for i in range(65, 68):
        dtype[i] = str

    df = pd.read_csv(os.path.join(path_folder, name_file),
                                        sep=";", 
                                        header=None, 
                                        dtype=dtype, 
                                        skiprows=1, 
                                        error_bad_lines=False, 
                                        encoding = "ISO-8859-1")

    df = df[df[12].notnull()]
    for day in DAYS.split(","):
        df["horaires_" + day] = df[START_NUMBER_COLUMN] + df[START_NUMBER_COLUMN + 1] + df[START_NUMBER_COLUMN + 2] + df[START_NUMBER_COLUMN + 3]
        df["horaires_" + day] = df["horaires_" + day].apply(lambda x: "{}:{}-{}:{} {}:{}-{}:{}".format(x[:2], x[2:4], x[4:6], x[6:8], x[8:10], x[10:12], x[12:14], x[14:]))
        for i in range(4):
            df.drop(START_NUMBER_COLUMN + i, axis=1, inplace=True)
        START_NUMBER_COLUMN += 4

    lib_col = {2: "code_point_relais", 3: "enseigne", 6: "latitude", 7: "longitude", 
              9: "adresse_1", 10: "adresse_2", 11: "adresse_3", 12: "code_postal", 
              13: "ville", 55: "debut_absence_1", 56: "fin_absence_1", 58: "debut_absence_2", 
              59: "fin_absence_2", 61: "debut_absence_3", 62: "fin_absence_3"}
    df.rename(columns=lib_col, inplace=True)

    NAME_DATE_COLUMNS = ["debut_absence_1", "fin_absence_1", "debut_absence_2", "fin_absence_2", "debut_absence_3", "fin_absence_3"]
    for name_date_column in NAME_DATE_COLUMNS:
        df[name_date_column] = df[name_date_column].apply(format_date)

    df["categorie_pr_chronopost"] = "C13"
    return df


def file_treatment_C9(path_folder, name_file):
    print("TRAITEMENT DU FICHIER C9: " + name_file)
    lib_col = {"Point Relais": "code_point_relais", "Enseigne": "enseigne", "Nom": "nom"}
    lib_col.update({"Adresse 1": "adresse_1", "Adresse 2": "adresse_2", "Adresse 3": "adresse_3"})
    lib_col.update({"Code Postal": "code_postal", "Horaires Lundi": "horaires_lundi"})
    lib_col.update({"Horaires Mardi": "horaires_mardi", "Horaires Mercredi": "horaires_mercredi"})
    lib_col.update({"Horaires Jeudi": "horaires_jeudi", "Horaires Vendredi": "horaires_vendredi"})
    lib_col.update({"Horaires Samedi": "horaires_samedi", "Horaires Dimanche": "horaires_dimanche"})
    lib_col.update({"Debut Absence": "debut_absence_1", "Fin Absence": "fin_absence_1"})
    lib_col.update({"Debut Absence.1": "debut_absence_2", "Fin Absence.1": "fin_absence_2"})
    lib_col.update({"Debut Absence.2": "debut_absence_3", "Fin Absence.2": "fin_absence_3", "Ville": "ville"})
    
    df = pd.read_csv(os.path.join(path_folder, name_file),
                                    sep=";",
                                    dtype={"Code Postal": str},
                                    error_bad_lines=False, 
                                    encoding = "ISO-8859-1")

    df = df[df["Ville"].notnull()]
    df.rename(columns=lib_col, inplace=True)
    df["categorie_pr_chronopost"] = "C9"
    return df


@time_execution("TRANSFORMATION DES FICHIERS CSV EN FICHIERS EXCEL")
def transform_csv_to_excel():
    list_csv_files = os.listdir(os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_CSV))
    list_excel_files = os.listdir(os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_EXCEL))
    list_excel_files = [file_excel.split(".")[0] for file_excel in list_excel_files]

    for name_file in list_csv_files:
        if not(name_file.split(".")[0] in list_excel_files):
            if "csv" in name_file and "C9" in name_file:
                file = file_treatment_C9(os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_CSV), name_file)
            elif "csv" in name_file and "C13" in name_file:
                file = file_treatment_C13(os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_CSV), name_file)
            file.to_excel(os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_EXCEL, name_file.split(".")[0] + ".xlsx"), index=False)


if __name__ == "__main__":
    transform_csv_to_excel()
