# imap_archive

Make an offline archive of an email account.

First install the requirements: 

```
python -m pip install -r requirements.txt
```

and then run:

```
python imap_email_archiver.py
```

Change the behavior by editing the variables `SAVE_FOLDER`, `SAVE_FILES` and `FOLDER_LIMIT` in `imap_email_archiver.py`.

## Credentials

`credentials.json`:

```json
{
    "username": "lyra.belacqua@gmail.com",
    "password": "pantala1m0n",
    "server_name": "imap.google.com"
}
```