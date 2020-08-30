#!/usr/bin/python3

# Standard library
import os, sys, json, re
from shutil import rmtree
from functools import partial

import shutil
shell_width = shutil.get_terminal_size((80, 20)).columns

# External packages
import imap_tools
from pathvalidate import sanitize_filename
from progressbar import ProgressBar # progressbar2 library

import email_to_json

def write_to_file(filename, payload, as_bytes):
    method = "wb" if as_bytes else "w"
    with open(filename, method) as write_file:
        write_file.write(payload)

def enumerate_file_path(path):
    """ Add '_#' to the path to make it unique """
    if os.path.exists(path):
        a = os.path.splitext(path)
        i = 1
        while True:
            enumerated_path = "{}_{}{}".format(a[0], i, a[1])
            if not os.path.exists(enumerated_path):
                return enumerated_path
            i += 1
    else:
        return path

# TODO Not pretty, needs improving, especially with the arguments
def write_out_html(subject, folder_name, body):
    subject = sanitize_filename(subject)
    filename = f"{subject[:50]}.html"
    file_path = enumerate_file_path(os.path.join(folder_name, filename))
    with open(file_path, "w") as f:
        f.write(body)

def make_folder_if_absent(path, folder_name):
    full_path = os.path.join(path, folder_name)
    if not os.path.isdir(full_path):
        os.mkdir(full_path)
    return full_path

def confirm(msg="Continue?"):
    answer = ""
    while answer not in ["y", "n"]:
        answer = input(f"{msg} [Y/N]?").lower()
    return answer == "y"

if __name__ == "__main__":

    SAVE_FOLDER = "emails"
    SAVE_FILES = True
    FOLDER_LIMIT = None

    # Clean up old ./emails directory
    if SAVE_FILES:
        save_location = os.path.join(".",SAVE_FOLDER)
        if os.path.isdir(save_location):
            if confirm("Overwrite?"):
                print(f"Cleaning up the ./{SAVE_FOLDER} directory")
            else:
                print("Aborted")
                sys.exit(1)
            rmtree(save_location)
        make_folder_if_absent(".",save_location)

    # Read credentials
    with open("credentials.json", "r") as read_file:
        credentials = json.load(read_file)
    username = credentials["username"]
    password = credentials["password"]
    server = credentials["server_name"]

    # Download messages from the mailbox
    with imap_tools.MailBox(server).login(username, password) as mailbox:
        for folder in mailbox.folder.list():
            folder_name = folder["name"]
            mailbox.folder.set(folder_name)
            fetch = partial(mailbox.fetch,reverse=False, mark_seen=False, limit=FOLDER_LIMIT)

            n_messages = mailbox.folder.status(folder_name)["MESSAGES"]
            print(f"Folder {folder_name} | Messages {n_messages}")
            if n_messages==0: continue # Skip if nothing there

            with ProgressBar(max_value=n_messages-1,prefix=f"{folder_name}:",redirect_stdout=True) as bar:
                for i, msg in enumerate(fetch(headers_only=(not SAVE_FILES))):
                    subject = msg.subject
                    imap_failed = subject[:41] == 'Retrieval using the IMAP4 protocol failed'
                    if imap_failed: # Pick subject from text
                        match = re.search('Subject: "(.+)"(?=[\r|$])',msg.text)
                        subject = match.groups()[0] if match else ''
                    if SAVE_FILES:
                        # Make folder for the mailbox
                        mailbox_folder = make_folder_if_absent(save_location,sanitize_filename(folder_name))

                        # Make folder for email thread with same name as msg.subject
                        sanitized_subject = sanitize_filename(subject)
                        if not sanitized_subject:
                            sanitized_subject = enumerate_file_path("No subject")
                        subject_folder = make_folder_if_absent(mailbox_folder, sanitized_subject)

                        json_filename = enumerate_file_path(os.path.join(subject_folder, "message.json"))

                        if imap_failed:
                            match = re.search('From: "(.+)"(?=[\r|$])',msg.text)
                            from_data = match.groups()[0] if match else ''
                            match = re.search('Sent date: "(.+)"(?=[\r|$])',msg.text)
                            sent_date = match.groups()[0] if match else ''
                            data = {'subject':subject, 'from':from_data, 'sent_date':sent_date}
                            write_to_file(json_filename, json.dumps(data,indent=4), as_bytes=False)
                        else:
                            encoded_message = email_to_json.json_encode(msg)
                            write_to_file(json_filename, encoded_message, as_bytes=False)

                            # To decode json representation of imap_tools.message.MailMessage object:
                            # b = email_to_json.json_decode(a)

                            write_out_html(subject, subject_folder, msg.html)
                            for att in msg.attachments:
                                file_path = enumerate_file_path(os.path.join(subject_folder, sanitize_filename(att.filename)))
                                write_to_file(file_path, att.payload, as_bytes=True)
                    # Print message number and subject
                    bar.update(i)
                    line = f"Msg {i} | " + subject.replace('\n','').replace('\r','')
                    if len(line)>shell_width:
                        line = line[:(shell_width-3)] + '...'
                    print(line)
    print("DONE")