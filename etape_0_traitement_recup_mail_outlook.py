import win32com.client
import os
from time import time

PATH_SORTIE = r"\\apps\Vol1\Data\011-BO_XI_entrees\07-DOR_DP\Sorties\FRANCIS\POINT RELAIS"
PATH_ONEDRIVE = r"d:\Users\Khaled-Khodja\OneDrive - TDF SAS\SCRIPTS\TRAITEMENT_MAIL_CHRONOPOST_PR\datas"
FOLDER_C9_C13_ORIGINALS = "0_C9_C13_ORIGINAUX"

EMAIL_INBOX = "referentiel_logistique"
FOLDER_1 = "Boîte de réception"
FOLDER_2 = "Courrier indésirable"
SENDER_EMAIL = ["chrexp@ediserv.chronopost.fr"]
SUBJECT = "Points relais CHRONOPOST pour TDF"

def time_execution(treatment_title):
    def decorator(func):
        def wrapper(*args, **kwargs):
            print("{0}\n{1}\n{0}".format("*" * len(treatment_title), treatment_title))
            t0 = time()
            func(*args, **kwargs)
            print("Durée du traitement: {} secondes".format(time() - t0))
            print("{0}\nFIN DE {1}\n{0}".format("*" * len(treatment_title), treatment_title))
        return wrapper
    return decorator

@time_execution("RECUPERATION DES FICHIERS POINT RELAIS DEPUIS OUTLOOK")
def recup_mail_chronopost():
    # Liste des fichiers existants dans le dossier de sauvegarde
    pudo_files = os.listdir(os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_ORIGINALS))

    # Choix de la boite mail à scruter
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Récupération des mails
    for folder in [FOLDER_1, FOLDER_2]:
        inbox = outlook.Folders(EMAIL_INBOX).Folders(folder)
        messages = inbox.items
        for message in messages:
            try:
                if message.SenderName in SENDER_EMAIL:
                    print(message.SenderName)
                    nombre_fichiers_joint = message.Attachments.Count
                    if nombre_fichiers_joint > 0:
                        for i in range(1, nombre_fichiers_joint + 1):
                            nom_fichier_joint = message.Attachments.item(i).FileName
                            print("Fichier joint trouvé: {}".format(nom_fichier_joint))
                            if not(nom_fichier_joint in pudo_files):
                                file_name_path = os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_ORIGINALS, nom_fichier_joint)
                                print("Sauvegarde du fichier: {}".format(nom_fichier_joint))
                                message.Attachments.item(i).SaveAsFile(file_name_path)
                    elif message.Subject == SUBJECT:
                        body = message.Body
                        received_time_date = message.ReceivedTime.strftime("%Y%m%d")

                        file_name = "CHRONO_RELAIS_C13_DETAILS_CHRONOS_" + received_time_date + ".csv"
                        print("Fichier dans corps du message trouvé: {}".format(file_name))
                        if not(file_name in pudo_files):
                            file_name_path = os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_ORIGINALS, file_name)
                            print("Sauvegarde du fichier dans le corps du message: {}".format(file_name))
                            with open(file_name_path, "w") as f:
                                body2 = body.split("\r\n")
                                for row in body2:
                                    f.write(row + "\n")
                                body2 = None
            except TypeError as e:
                print("Probleme detecte: ", e)
                continue


if __name__ == "__main__":
    recup_mail_chronopost()
