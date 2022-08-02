import os
import shutil
import win32com.client as win32
from pathlib import Path


def create_language_list(list_of_languages):
    language_file = open("Source/language_list.txt", "r")
    for language in language_file:
        list_of_languages.append(language.strip("\n"))


def move_files(source_path, upload_path):
    # existing_files = []
    try:
        for file in os.listdir(source_path):
            full_target = upload_path + file
            check_file = Path(full_target)
            full_source = source_path + file
            if not check_file.is_file() and full_target.endswith(".pdf"):
                shutil.copyfile(full_source, full_target)
            # existing_files.append(f'[{datetime.now()}]\t{check_file}')
    except Exception as e:
        print("Error while moving file to destination:" + str(e))
    # write_log_file(existing_files, upload_path)


def write_log_file(existing_files, upload_path):
    try:
        log_file_existing = open(fr'{upload_path}Existing_Files.txt', "a")
        for entry in existing_files:
            log_file_existing.writelines(entry)
        log_file_existing.close()
    except Exception as e:
        print("Error while writing log-file:" + str(e))


def rename_source_files(upload_path):
    special_char_map = {ord('ä'): 'ae', ord('Ä'): 'Ae', ord('ü'): 'ue', ord('Ü'): 'Ue', ord('ö'): 'oe', ord('Ö'): 'Oe',
                        ord('ß'): 'ss', ord('+'): '_', ord('&'): '', ord("'"): '', ord('á'): 'a', ord('é'): 'e',
                        ord('à'): 'a',  ord('ú'): 'u', ord('ô'): 'o', ord('ó'): 'o', ord('ò'): 'o', ord('–'): '',
                        ord('â'): 'a',  ord('î'): 'i', ord('’'): '', ord('è'): 'e'}
    try:
        for file in os.listdir(upload_path):
            renamed_file = file.translate(special_char_map)
            if renamed_file.endswith(" .pdf"):
                posn = len(renamed_file)-5
                renamed_file = renamed_file[:posn] + renamed_file[-4:]
            renamed_file = renamed_file.replace(" ", "_")
            renamed_file = renamed_file.replace("-", "_")
            os.rename(f'{upload_path}{file}', f'{upload_path}{renamed_file}')
    except Exception as e:
        print("Error while renaming files:" + str(e))


def write_sti_file(upload_path, list_of_languages):
    try:
        sti_target_file = open(f'{upload_path}document_upload_file.sti', "w")
        sti_target_file.writelines("<stimport>\n")
        list_of_files = []

        for file in os.listdir(upload_path):
            if file.endswith(".pdf"):
                list_of_files.append(file.split("_")[:-1])
                language = file.split("_")[-1].split(".")[0].strip(" ").lower()
                if file.split("_")[:-1] in list_of_files and len(file.split("_")) > 3 and language in list_of_languages:
                    sti_target_file.writelines(f'<node id="{file[:12]} class="/class::BaseNodeClass/'
                                               f'DocuManagerBaseClass/Resource/DocumentResource" '
                                               f'parent="5393570059">\n')
                    sti_target_file.writelines(f'<attribute name="Resource" type="resource" aspect="{language}">'
                                               f'{upload_path}{file}</attribute>\n')
                    sti_target_file.writelines(f'<attribute name="Title" type="string" aspect="{language}">'
                                               f'{file}</attribute>\n')
                else:
                    sti_target_file.writelines(f'<node id="{file}" class="/class::BaseNodeClass/DocuManagerBaseClass'
                                               f'/Resource/DocumentResource" parent="5393570059">\n')
                    sti_target_file.writelines(f'<attribute name="Resource" type="resource" aspect="de">'
                                               f'{upload_path}{file}</attribute>\n')
                    sti_target_file.writelines(f'<attribute name="Title" type="string" aspect="de">'
                                               f'{file}</attribute>\n')
                sti_target_file.writelines(f'</node>\n')
        sti_target_file.writelines("</stimport>\n")
        sti_target_file.close()
    except Exception as e:
        print("Error while creating STI file:" + str(e))


def send_notification(upload_path, fileName="document_upload_file.sti"):
    outlook = win32.Dispatch('outlook.application')
    new_mail = outlook.CreateItem(0)
    mailing_list = open("Source/email-recepients.txt", "r")
    recipient = ""
    for i in mailing_list:
        recipient = recipient + i.strip("\n")
    new_mail.To = recipient
    new_mail.Subject = "Neue ST4-Importdatei"
    new_mail.Body = "Hallo,\n\n" \
                    "es ist eine STI-Datei erzeugt worden.\n" \
                    "Angehängt ist die Datei mit den zu importierenden Einträgen\n" \
                    "Die Datei kann über NotePad geöffnet und verändert werden.\n\n" \
                    "Viele Grüße\n" \
                    "Andreas Neumann"
    mail_attachment = f'{upload_path}{fileName}'
    new_mail.Attachments.Add(mail_attachment)
    new_mail.Send()


language_list = []
source_path = "P:\\Engineering\\FB_tech_Doku\\ET-Listen\\Druckdateien Maschinen Januar 2003\\01 Printkataloge\\"
upload_path = "S:\\00_Admin\\InfoCube\\Printkataloge\\"

if __name__ == "__main__":

    create_language_list(language_list)
    move_files(source_path, upload_path)
    rename_source_files(upload_path)
    write_sti_file(upload_path, source_path)
    # send_notification(upload_path)
