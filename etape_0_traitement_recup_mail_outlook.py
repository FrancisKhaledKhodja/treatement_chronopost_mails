import win32com.client
import os
import time

PATH_SORTIE = r"\\apps\Vol1\Data\011-BO_XI_entrees\07-DOR_DP\Sorties\FRANCIS\POINT RELAIS"
PATH_ONEDRIVE = r"d:\Users\Khaled-Khodja\OneDrive - TDF SAS\SCRIPTS\TRAITEMENT_MAIL_CHRONOPOST_PR\datas"
FOLDER_C9_C13_ORIGINALS = "0_C9_C13_ORIGINAUX"

EMAIL_INBOX = "referentiel_logistique"
FOLDER_1 = "Boîte de réception"
FOLDER_2 = "Courrier indésirable"
SENDER_EMAIL = ["chrexp@ediserv.chronopost.fr"]
SUBJECT = "Points relais CHRONOPOST pour TDF"

def recup_mail_chronopost():
    t0 = time.time()
    TREATMENT_TITLE = "RECUPERATION DES FICHIERS POINT RELAIS DEPUIS OUTLOOK"
    print("{0}\n{1}\n{0}".format("*" * len(TREATMENT_TITLE), TREATMENT_TITLE))

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
                                chemin_nom_fichier = os.path.join(PATH_ONEDRIVE, rep_c9_c13_originals, nom_fichier_joint)
                                print("Sauvegarde du fichier: {}".format(nom_fichier_joint))
                                message.Attachments.item(i).SaveAsFile(chemin_nom_fichier)
                    elif message.Subject == SUBJECT:
                        body = message.Body
                        date_reception = message.ReceivedTime.strftime("%Y%m%d")

                        file_name = "CHRONO_RELAIS_C13_DETAILS_CHRONOS_" + date_reception + ".csv"
                        print("Fichier dans corps du message trouvé: {}".format(file_name))
                        if not(file_name in pudo_files):
                            chemin_nom_fichier = os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_ORIGINALS, file_name)
                            print("Sauvegarde du fichier dans le corps du message: {}".format(file_name))
                            with open(chemin_nom_fichier, "w") as f:
                                body2 = body.split("\r\n")
                                for row in body2:
                                    f.write(row + "\n")
                                body2 = None
            except TypeError as e:
                print("Probleme detecte: ", e)
                continue

    TREATMENT_TITLE = "FIN DE LA RECUPERATION DES FICHIERS POINT RELAIS DEPUIS OUTLOOK"
    print("Durée du traitement: {} secondes".format(time.time() - t0))
    print("{0}\n{1}\n{0}".format("*" * len(TREATMENT_TITLE), TREATMENT_TITLE))


if __name__ == "__main__":
    recup_mail_chronopost()
