import win32com.client
import os
from time import time

from constants import *


def time_execution(treatment_title: str):
    """decorator measuring the execution time of a function

    Args:
        treatment_title (str): Name of the function, for instance.
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            print("{0}\n{1}\n{0}\n".format("*" * len(treatment_title), treatment_title))
            t0 = time()
            func(*args, **kwargs)
            print("Durée du traitement: {} secondes".format(time() - t0))
            print("{0}\nFIN DE {1}\n{0}\n".format("*" * len(treatment_title), treatment_title))
        return wrapper
    return decorator


def file_attachment_recovery(email_inbox: str, sender_email: list, list_folders_outlook: list, backup_recovery: str, list_files_already_saved: list):
    """Save attached files in an outlook email

    Args:
        email_inbox (str): inbox to scan
        sender_email (list): list of emails to scan
        list_folders_outlook (list): list of folders to scan
        backup_recovery (str): absolute path to save the files
        list_files_already_saved (list): list of files that you do not want to save again
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    for folder in list_folders_outlook:
        inbox = outlook.Folders(email_inbox).Folders(folder)
        messages = inbox.items
        for message in messages:
            try:
                if message.SenderName in sender_email:
                    files_attachment_number = message.Attachments.Count
                    if files_attachment_number > 0:
                        for i in range(1, files_attachment_number + 1):
                            file_attachment_name = message.Attachments.item(i).FileName
                            print("Fichier joint trouvé: {}".format(file_attachment_name))
                            if file_attachment_name not in list_files_already_saved:
                                file_name_path = os.path.join(backup_recovery, file_attachment_name)
                                print("Sauvegarde du fichier: {}".format(file_attachment_name))
                                message.Attachments.item(i).SaveAsFile(file_name_path)
            except TypeError as e:
                print("Probleme detecte: ", e)
                continue


def body_email_recovery(email_inbox: str, sender_email: list, list_folders_outlook: list, subject: str, backup_recovery: str, file_name_header, list_files_already_saved: list):
    """Save body from an outlook email

    Args:
        email_inbox (str): inbox to scan
        sender_email (list): emails to scan
        list_folders_outlook (list): list of folders to scan
        subject (str): email's subject to select
        backup_recovery (str): absolute path to save the files
        file_name_header ([type]): file name header given to the email's body
        list_files_already_saved (list): list of files that you do not want to save again
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    for folder in list_folders_outlook:
        inbox = outlook.Folders(email_inbox).Folders(folder)
        messages = inbox.items
        for message in messages:
            try:
                if message.SenderName in sender_email:
                    if message.Subject == subject:
                        body = message.Body
                        received_time_date = message.ReceivedTime.strftime("%Y%m%d")
                        file_name = f"{file_name_header}_{received_time_date}.csv"
                        print("Fichier dans corps du message trouvé: {}".format(file_name))
                        if file_name not in list_files_already_saved:
                            file_name_path = os.path.join(backup_recovery, file_name)
                            print("Sauvegarde du fichier dans le corps du message: {}".format(file_name))
                            with open(file_name_path, "w") as f:
                                body = body.split("\r\n")
                                for row in body:
                                    f.write(row + "\n")
            except TypeError as e:
                print("Probleme detecte: ", e)
                continue


@time_execution("RECUPERATION DES FICHIERS POINT RELAIS DEPUIS OUTLOOK")
def recup_mail_chronopost():
    pudo_files = os.listdir(os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_ORIGINALS))
    file_attachment_recovery(EMAIL_INBOX, SENDER_EMAIL, [FOLDER_1, FOLDER_2], os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_ORIGINALS), pudo_files)
    body_email_recovery(EMAIL_INBOX, SENDER_EMAIL, [FOLDER_1, FOLDER_2], SUBJECT, os.path.join(PATH_ONEDRIVE, FOLDER_C9_C13_ORIGINALS), FILE_NAME_HEADER_C13, pudo_files)


if __name__ == "__main__":
    recup_mail_chronopost()
