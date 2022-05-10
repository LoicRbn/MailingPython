
import imaplib
import email
from email.header import decode_header
from getpass import getpass


# outlook.office365.com
imap_adress = input("Veuillez saisir l'adresse imap de votre service de messagerie : ")
imap = imaplib.IMAP4_SSL(imap_adress)
#on demande le login
login = input("Saisir votre login : ")
#on demande le mdp avec une fonction plus sécurisé que le input pour ne pas voir le mdp saisit
passwd = getpass("Saisir votre mot de passe : ")
imap.login(login, passwd)

status, messages = imap.select("INBOX")
# top N premiers mails à récupérer
N = 10
# nombre de total de mails
messages = int(messages[0])

j = 1
for i in range(messages, messages-N, -1):
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
            subject, encoding = decode_header(msg["Subject"])[0]
            #gestion du cas ou le sujet est vide
            if encoding != None:
                if isinstance(subject, bytes):
                    subject = subject.decode(str(encoding))

            From, encoding = decode_header(msg.get("From"))[0]
            if isinstance(From, bytes):
                From = From.decode(encoding)

            print("Mail numéro ", j)
            print("Subject:", subject)
            print("From:", From)
    if j == 1:
        stock_subject = str(subject)
    j += 1
# on ferme la connexion
imap.close()
imap.logout()

print("Sujet du premier mail", stock_subject)

import smtplib


# adresse email a qui envoyer
receiver = input("Saisir l'adresse mail du destinataire : ")
# smtp.office365.com

mailserver = smtplib.SMTP('SMTP.office365.com', 587)
mailserver.connect('SMTP.office365.com', 587)
mailserver.ehlo()
mailserver.starttls()
mailserver.login(login, passwd)
mailserver.sendmail(login, receiver, stock_subject)
mailserver.quit()