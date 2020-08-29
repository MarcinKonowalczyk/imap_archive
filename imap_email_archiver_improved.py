import imap_tools
import json
from pathvalidate import sanitize_filename
import os
import email_to_json


def plural(num):
    try:
        if num > 1:
            return 's'
        else:
            return ''
    except TypeError:
        pass


def write_to_file(filename, payload, as_bytes):
    if as_bytes:
        method = 'wb'
    else:
        method = 'w'
    with open(filename, method) as write_file:
        write_file.write(payload)


def enumerate_file_path(path):
    if os.path.exists(path):
        a = os.path.splitext(path)
        i = 1
        while True:
            enumerated_path = '{}_{}{}'.format(a[0], i, a[1])
            if not os.path.exists(enumerated_path):
                return enumerated_path
            i += 1
    else:
        return path


def write_out_html(subject, folder_name, body):
    """Not pretty, needs improving, especially with the arguments"""
    subject = sanitize_filename(subject)
    filename = f"{subject[:50]}.html"
    file_path = enumerate_file_path(os.path.join(folder_name, filename))
    open(file_path, "w").write(body)


def make_folder_if_absent(path, folder_name):
    full_path = os.path.join(path, folder_name)
    if not os.path.isdir(full_path):
        os.mkdir(full_path)
    return full_path

if __name__ == "__main__":

    SAVE_LOCATION = "./emails/"
    SAVE_FILES = True
    FOLDER_LIMIT = None

    with open("credentials.json", "r") as read_file:
        credentials = json.load(read_file)

    username = credentials["username"]
    password = credentials["password"]
    server = credentials["server_name"]

    with imap_tools.MailBox(server).login(username, password) as mailbox:
        for folder in mailbox.folder.list():
            mailbox_name = folder["name"]
            mailbox.folder.set(mailbox_name)
            mailbox_folder = make_folder_if_absent(SAVE_LOCATION, sanitize_filename(mailbox_name))
            for i, msg in enumerate(mailbox.fetch(reverse=False, mark_seen=False, limit=FOLDER_LIMIT)):

                if SAVE_FILES:
                    # Make folder for email thread with same name as msg.subject
                    sanitized_subject = sanitize_filename(msg.subject)
                    if sanitized_subject == '':
                        sanitized_subject = enumerate_file_path('No subject')
                    subject_folder = make_folder_if_absent(mailbox_folder, sanitized_subject)

                    json_filename = enumerate_file_path(os.path.join(subject_folder, 'message.json'))
                    encoded_message = email_to_json.json_encode(msg)
                    write_to_file(json_filename, encoded_message, as_bytes=False)

                    # To decode json representation of imap_tools.message.MailMessage object:
                    # b = email_to_json.json_decode(a)

                    write_out_html(msg.subject, subject_folder, msg.html)
                    for att in msg.attachments:
                        file_path = enumerate_file_path(os.path.join(subject_folder, att.filename))
                        write_to_file(file_path, att.payload, as_bytes=True)

                print('\r' * len(str(i - 1)),
                    f'{mailbox_name}: downloaded {i} message{plural(i)}...', end='',flush=True)