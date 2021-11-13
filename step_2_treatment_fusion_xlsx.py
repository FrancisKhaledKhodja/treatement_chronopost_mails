
import os
import pandas as pd

from constants import *
from various_functions import time_execution


def fusion_file_xlsx_c9_c13(path_file: str, name_file_1: str, name_file_2: str):
    col = ["code_point_relais", "enseigne", "nom", "adresse_1", "adresse_2"]
    col.extend(["adresse_3", "code_postal", "ville", "horaires_lundi", "horaires_mardi", "horaires_mercredi"])
    col.extend(["horaires_jeudi", "horaires_vendredi", "horaires_samedi", "horaires_dimanche"])
    col.extend(["debut_absence_1", "fin_absence_1", "debut_absence_2", "fin_absence_2", "debut_absence_3"])
    col.extend(["fin_absence_3", "categorie_pr_chronopost", "latitude", "longitude"])
    
    if "C9" in name_file_1:
        df_c9 = pd.read_excel(os.path.join(path_file, name_file_1), dtype={"code_postal": str})
        df_c13 = pd.read_excel(os.path.join(path_file, name_file_2), dtype={"code_postal": str})
    else:
        df_c9 = pd.read_excel(os.path.join(path_file, name_file_2), dtype={"code_postal": str})
        df_c13 = pd.read_excel(os.path.join(path_file, name_file_1), dtype={"code_postal": str})
            
    code_pr_c9 = set(df_c9["code_point_relais"].tolist())
    code_pr_c13 = set(df_c13["code_point_relais"].tolist())
    # Identification des groupes C9-C13, des purs C9 et des purs C13
    code_pr_c9_c13 = code_pr_c9.intersection(code_pr_c13)
    code_pr_c9_without_c13 = code_pr_c9.difference(code_pr_c13)
    code_pr_c13_without_c9 = code_pr_c13.difference(code_pr_c9)
    # Construction de la nouvelle table des points relais
    df_c9_c13 = df_c9[df_c9["code_point_relais"].isin(list(code_pr_c9_c13))]
    df_c9_c13["categorie_pr_chronopost"] = "C9_C13"
    df_c9_without_c13 = df_c9[df_c9["code_point_relais"].isin(list(code_pr_c9_without_c13))]
    df_c13_without_c9 = df_c13[df_c13["code_point_relais"].isin(list(code_pr_c13_without_c9))]
    df = pd.concat([df_c9_c13, df_c9_without_c13, df_c13_without_c9], sort=True)
    df = df[col]
    df.sort_values("code_point_relais", inplace=True)

    return df


@time_execution("RASSEMPLEMENT DES FICHIERS C9 ET C13")
def treatment_fusion_xlsx_c9_c13(path_folder_excel, path_folder_fusion_excel):
    list_files_fusion_excel = os.listdir(path_folder_fusion_excel)
    list_files_excel = os.listdir(path_folder_excel)

    dates_fichiers = sorted(list(set([x.split(".")[0][-8:] for x in list_files_excel])))
    dates_fichiers.reverse()

    for date in dates_fichiers:
        list_files_with_date_selected = list(filter(lambda x: date in x, list_files_excel))
        if len(list_files_with_date_selected) == 2:
            list_file_with_date_selected =  list(filter(lambda x: date in x, list_files_fusion_excel))
            if len(list_file_with_date_selected) == 0:
                print("Traitement des fichiers du: " + date)
                df = fusion_file_xlsx_c9_c13(path_folder_excel, list_files_with_date_selected[0], list_files_with_date_selected[1])
                df.to_excel(os.path.join(path_folder_fusion_excel, "CHRONO_RELAIS_C9_C13_DETAILS_CHRONOS_{}.xlsx".format(date)),index=False)

    
if __name__ == "__main__":
    treatment_fusion_xlsx_c9_c13(os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_EXCEL), os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_FUSION_EXCEL))
