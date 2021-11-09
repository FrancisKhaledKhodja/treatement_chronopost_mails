import win32com.client
import os
import time

path_sortie = r"\\apps\Vol1\Data\011-BO_XI_entrees\07-DOR_DP\Sorties\FRANCIS\POINT RELAIS"
path_onedrive = r"d:\Users\Khaled-Khodja\OneDrive - TDF SAS\SCRIPTS\TRAITEMENT_MAIL_CHRONOPOST_PR\datas"
rep_c9_c13_originals = "0_C9_C13_ORIGINAUX"

boite_mail_choisie = "referentiel_logistique"
folder_1 = "Boîte de réception"
folder_2 = "Courrier indésirable"
expediteur_a_surveiller = ["chrexp@ediserv.chronopost.fr"]
objet = "Points relais CHRONOPOST pour TDF"


def recup_mail_chronopost():
    t0 = time.time()
    titre_traitement = "RECUPERATION DES FICHIERS POINT RELAIS DEPUIS OUTLOOK"
    print("{0}\n{1}\n{0}".format("*" * len(titre_traitement), titre_traitement))

    # Liste des fichiers existants dans le dossier de sauvegarde
    fichiers_pudo = os.listdir(os.path.join(path_onedrive, rep_c9_c13_originals))

    # Choix de la boite mail à scruter
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Récupération des mails
    for folder in [folder_1, folder_2]:
        inbox = outlook.Folders(boite_mail_choisie).Folders(folder)
        messages = inbox.items
        for message in messages:
            try:
                if message.SenderName in expediteur_a_surveiller:
                    print(message.SenderName)
                    nombre_fichiers_joint = message.Attachments.Count
                    if nombre_fichiers_joint > 0:
                        for i in range(1, nombre_fichiers_joint + 1):
                            nom_fichier_joint = message.Attachments.item(i).FileName
                            print("Fichier joint trouvé: {}".format(nom_fichier_joint))
                            if not(nom_fichier_joint in fichiers_pudo):
                                chemin_nom_fichier = os.path.join(path_onedrive, rep_c9_c13_originals, nom_fichier_joint)
                                print("Sauvegarde du fichier: {}".format(nom_fichier_joint))
                                message.Attachments.item(i).SaveAsFile(chemin_nom_fichier)
                    elif message.Subject == objet:
                        body = message.Body
                        date_reception = message.ReceivedTime.strftime("%Y%m%d")

                        file_name = "CHRONO_RELAIS_C13_DETAILS_CHRONOS_" + date_reception + ".csv"
                        print("Fichier dans corps du message trouvé: {}".format(file_name))
                        if not(file_name in fichiers_pudo):
                            chemin_nom_fichier = os.path.join(path_onedrive, rep_c9_c13_originals, file_name)
                            print("Sauvegarde du fichier dans le corps du message: {}".format(file_name))
                            with open(chemin_nom_fichier, "w") as f:
                                body2 = body.split("\r\n")
                                for row in body2:
                                    f.write(row + "\n")
                                body2 = None
            except AttributeError:
                print("Probleme detecte")
                continue

    titre_traitement = "FIN DE LA RECUPERATION DES FICHIERS POINT RELAIS DEPUIS OUTLOOK"
    print("Durée du traitement: {} secondes".format(time.time() - t0))
    print("{0}\n{1}\n{0}".format("*" * len(titre_traitement), titre_traitement))


if __name__ == "__main__":
    recup_mail_chronopost()
