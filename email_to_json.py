import json
import datetime

import imap_tools


def attachment_to_dict(attachment):
    return {"filename": attachment.filename,
            "content_type": attachment.content_type  # ,
            # "payload": attachment.payload #enabling this causes byte conversion issues
            }


def json_encode(data):
    return EmailJSONEncoder(indent=4).encode(data)


def json_decode(string):
    return EmailJSONDecoder().decode(string)


def to_iso(date):
    return datetime.datetime.strftime(date, '%a, %d %b %Y %H:%M:%S %z')


def from_iso(date_string):
    return datetime.datetime.strptime(date_string, '%a, %d %b %Y %H:%M:%S %z')


class EmailJSONEncoder(json.JSONEncoder):

    def default(self, obj):

        if isinstance(obj, imap_tools.message.MailMessage):
            # consider wrapping each of these in a try/catch decorator?
            return {"__imap_tools.message.MailMessage__": True,
                    # "uid": obj.uid, # serialising UID not recommended
                    "subject": obj.subject,
                    "date": obj.date,
                    "from": obj.from_,
                    "to": obj.to,
                    "cc": obj.cc,
                    "bcc": obj.bcc,
                    "reply_to": obj.reply_to,
                    "text": obj.text,
                    "html": obj.html,
                    "flags": obj.flags,
                    "from_values": obj.from_values,
                    "to_values": obj.to_values,
                    "cc_values": obj.cc_values,
                    "bcc_values": obj.bcc_values,
                    "reply_to_values": obj.reply_to_values,
                    "headers": obj.headers,
                    "attachments": [attachment_to_dict(att) for att in obj.attachments],
                    # "obj_to_string": obj.obj.as_string(), # not essential data
                    "_raw_uid_data": obj._raw_uid_data,
                    "_raw_flag_data": obj._raw_flag_data
                    }

        elif isinstance(obj, datetime.datetime):
            return {"__datetime__": True,
                    "string": to_iso(obj)}
        elif isinstance(obj, bytes):
            try:
                return str(obj, 'utf-8')
            except:
                print('\n utf-8 character set was not usable. Trying cp437...\n')
                return str(obj, 'cp437')
        return json.JSONEncoder(indent=4).default(obj)


class EmailJSONDecoder(json.JSONDecoder):
    def __init__(self, *args, **kwargs):
        json.JSONDecoder.__init__(self, object_hook=self.object_hook, *args, **kwargs)

    def object_hook(self, obj):

        if isinstance(obj, dict):
            if "__imap_tools.message.MailMessage__" in obj:
                raw_message_data = bytes(obj["obj_to_string"], 'utf-8')

                # when serialising raw_uid_data and raw_flag_data enabled in JSONEncoder
                raw_uid_data = bytes(obj["_raw_uid_data"], 'utf-8')
                raw_flag_data = [bytes(item, 'utf-8') for item in obj["_raw_flag_data"]]

                # else, keep the following code uncommented
                # raw_uid_data = b''
                # raw_flag_data = b''
                return imap_tools.MailMessage(((raw_uid_data, raw_message_data), raw_flag_data))
            elif "__datetime__" in obj:
                return from_iso(obj["string"])

        if isinstance(obj, dict):
            for key in list(obj):
                obj[key] = self.object_hook(obj[key])
            return obj

        if isinstance(obj, list):
            for i in range(0, len(obj)):
                obj[i] = self.object_hook(obj[i])
            return obj
        return obj
