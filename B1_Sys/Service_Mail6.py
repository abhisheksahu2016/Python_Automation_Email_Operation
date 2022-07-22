"""Create a connection to Gmail and do something with the results

References:
http://www.voidynullness.net/blog/2013/07/25/gmail-email-with-python-via-imap/
and
https://yuji.wordpress.com/2011/06/22/python-imaplib-imap-example-with-gmail/
"""
__author__ = 'abought'

import email.parser
import imaplib
import getpass
import sys
import re
import ssl

from pprint import pprint as pp

# User may want to change these parameters if running script as-is

# Search folders, multiple directories can be given
# TODO: A user will want to change this
SEARCH_FOLDER = ['"[Gmail]Trash"', '"[Gmail]/All Mail"', '"INBOX"'] 
DEFAULT_MAIL_SERVER = 'imap.gmail.com'

# No user parameters below this line
ADDR_PATTERN = re.compile("<(.+)>")  # Finds email as <nospam@nospam.com>


def connect(user, pwd, server=DEFAULT_MAIL_SERVER):
    """Connect to [the specified] mail server. Return an open connection"""
    conn = imaplib.IMAP4_SSL(host=server,
            ssl_context=ssl.create_default_context())
    try:
        conn.login(user, pwd)
    except imaplib.IMAP4.error:
        print("Failed to login")
        sys.exit(1)
    return conn


def print_folders(conn):
    """Print a list of open mailbox folders"""
    for f in conn.list():
        for i in f:
            print("\t", i)


def get_mails_from_folder(conn, folder_name):
    """Fetch a specific folder (or label) from server"""
    typ, data = conn.select(mailbox=folder_name, readonly=True)
    if typ != 'OK':
        print("Could not open specified folder. Known labels:")
        print_folders(conn)
        return

    typ, data = conn.search(None, 'ALL')
    if typ != 'OK':
        print("Could not get mail list of folder: ", folder_name)
        return

    return data[0].split()

def fetch_message(conn, msg_uid):
    """
    Fetch a specific message uid (not sequential id!) from the given folder;
    return the parsed message. User must ensure that specified
    message ID exists in that folder.
    """
    # TODO: Could we fetch just the envelope of the response to save bandwidth?
    typ, data = conn.fetch(msg_uid, '(RFC822)')

    if typ != 'OK':
        print("ERROR fetching message #", msg_uid)
        return

    return email.parser.BytesParser().parsebytes(data[0][1], headersonly=True)


def get_recipients(msg):
    """Given a parsed message, extract and return recipient list"""
    recipients = []
    addr_fields = ['From', 'To', 'Cc', 'Bcc', 'Reply-To', 'Sender']

    for f in addr_fields:
        if msg[f] is None:
            continue

        # str conversion is needed for non-ascii chars
        rlist = ADDR_PATTERN.findall(str(msg[f]))
        recipients.extend(rlist)

    return recipients


if __name__ == "__main__":
    # username = input("Enter username: ")
    # password = input("Enter password: ")

    username = "abhisheksakti2014@gmail.com"
    password = "rrgzirgsqkgzheuw"

    # Connect
    mail_conn = connect(username, password)

    # show folders of mail account
    #print_folders(mail_conn)

    # Open folders and get list of email message uids
    all_recipients = []
    for folder in SEARCH_FOLDER:
        # switch to folder
        mail_id_list = get_mails_from_folder(mail_conn, folder)
        if mail_id_list!=None:       
            for mail_id in get_mails_from_folder(mail_conn, folder):
                data = fetch_message(mail_conn, mail_id)
                recip_list = get_recipients(data)
                all_recipients.extend(recip_list)

    mail_conn.close()
    mail_conn.logout()

    # Very unsophisticated way of showing the recipient list
    print("List of all recipients:")
    print("------------")
    pp(all_recipients)

    print("\n\n List of all UNIQUE recipients:")
    print("-------------------------------")
    pp(set(all_recipients))
