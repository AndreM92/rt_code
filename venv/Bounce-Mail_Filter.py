import pandas as pd
import os
import re
from datetime import datetime

path = r"C:\Users\andre\OneDrive\Desktop\Marketing\ToDo/"
source_file = 'messages' + '.csv'
branch = 'Automatisierungstechnik'
########################################################################################################################
# Vorbereitung:
#1. Thunderbird öffnen → Menü → Erweiterungen/Add‑on und Themes
#2. Suche nach "ImportExportTools NG" -> installieren -> Thunderbird neu starten
#3. Nachrichten markieren -> Nachrichten exportieren im... -> CSV-Format -> Ordner auswählen
########################################################################################################################

# Extract text from elements
def extract_text(element):
    if element:
        if not isinstance(element, (str, int, float)):
            try:
                element = element.text.strip()
            except:
                return None
        element = str(element)
        if element == '':
            return element
        elif len(element) >= 1:
            repl_element = element.replace('\u200b', '').replace('\xa0', ' ').replace('\\xa0', ' ').replace(
                '\n', ' ')
            new_element = re.sub(r'\s+', ' ', repl_element).strip()
            return new_element
        else:
            return element

##No‑Such‑User / 550 5.1.1 — spezielle Hard‑Bounce‑Form, wenn die Empfängeradresse nicht existiert.
# 553 5.1.2 / 5.1.3 — Domain oder Adresse ungültig.
# 5.0.x / 5.1.x — permanente Fehlerklasse (nicht zustellbar).
# 554 5.4.12 – „Recipient address rejected: Access denied“ Das heißt: Die Nachricht konnte nicht zugestellt werden,
# weil der Server den Empfänger nicht akzeptiert hat.

def identify_mail(mail):
    mail_type = 'other'
    autoreply_keys = ['Automatische Antwort', 'Autoreply', 'AUTOREPLY', 'autoreply', 'Anfrage', 'Re: ', 'Abwesenheitsnotiz',
                      'Out of Office', 'Eingangsbestätigung', 'bin im Urlaub', 'Ticket', 'Anliegen bearbeiten',
                      'Anliegen umgehend bearbeiten', 'bei Ihnen melden', 'Vielen Dank', 'Automatic reply',
                      'Hope you’re', 'Thank you for', 'Request received', 'eingegangen',
                      'Nachricht', 'Guten Tag', 'received', 'Automatisierte']
    invalid_keys = ['550 5.', '551 5.', '552 5.', '553 5.', '554 5.', '555 5.', '556 5.',
                    '5.0.0', '5.1.1', '5.1.2', '5.1.3', '5.1.4.', 'Status: 5.0', 'ungültige E-Mail']
    bounce_keys = ['Unzustellbar', 'Returned mail', 'Returned Mail', 'Undeliverable', 'Undelivered', 'Delivery Status',
                   'Delivery Failure', 'Delivery failed', 'Achtung: ']
    for e in mail:
        e = extract_text(e)
        if not e or len(e) <= 4:
            continue
        if any(i in e for i in invalid_keys):
            mail_type = 'invalid_mail'
        if any(k in e for k in autoreply_keys) and not any(b in e for b in bounce_keys):
            mail_type = 'autoreply'
    return mail_type

def mail_to_list(temp_list, m_a):
    to_replace = ['mailto:', '(', ')', ':550', 'to:', '<', '>', ':', '...', '"']
    for r in to_replace:
        m_a = m_a.replace(r, '').strip().lower()
    if ';' in m_a:
        mail_parts = m_a.split(';')
        if len(mail_parts[0]) > 10:
            m_a = mail_parts[0]
        else:
            m_a = mail_parts[1]
    if m_a[-1] == '.' or m_a[-1] == ',':
        m_a = m_a[:-1]
    if 'mail' in m_a:
        ml = m_a.split('mail')
        if len(ml) == 2 and ml[0].strip() == ml[1].strip():
            m_a = ml[0].strip()
    if not '@research-tools' in m_a and not 'mailer-daemon@' in m_a and not 'postmaster@' in m_a \
            and not 'andre.muth@r' in m_a and not 'uwe.matzner@r' in m_a and not 'header.' in m_a \
            and not 'research-tools.net' in m_a and not 'redir+' in m_a \
            and (70 > len(m_a) > 10) \
            and m_a not in temp_list:
        temp_list.append(m_a)
    return temp_list

def analyze_mail(mail):
    temp_list = []
    full_text = ''
    subject = False
    for e in mail:
        e = extract_text(e)
        full_text = (full_text + ' ' + e).strip()
        if not e or len(e) <= 4:
            continue
        if 'Betreff:' in e:
            subject = e.replace('Betreff:','').strip()
        content_list = e.split(' ')
        for i in content_list:
            if '@' in i:
                temp_list = mail_to_list(temp_list, i)

    if not subject or len(temp_list) == 0:
        subject = full_text
    if len(temp_list) == 0:
        return '', '', subject

    user_mail = temp_list.pop(0)
    if not '.' in user_mail.split('@')[1]:
        user_mail = temp_list.pop(0)
    for t in temp_list:
        if t in user_mail:
            temp_list.remove(t)
    other_mails = temp_list
    if len(temp_list) == 0:
        return user_mail, '', subject
    if len(temp_list) == 1:
        return user_mail, temp_list[0], subject
    other_mails = str(temp_list).replace("'","").replace("[","").replace("]","")
    return user_mail, other_mails, subject


if __name__ == '__main__':
    os.chdir(path)
    returned_mails = pd.read_csv(source_file, header=None)
    table = []

    for id, mail in returned_mails.iterrows():
        invalid_mail = False
        autoreply = False
        mail_type = identify_mail(mail)

        user_mail, other_mails, subject = analyze_mail(mail)
        mail_data = [mail_type, user_mail, other_mails, subject]
        table.append(mail_data)
        print(mail_data)

    df_analyzed_mails = pd.DataFrame(table, columns=['mail_type', 'user_mail', 'other_mails', 'content'])
    # Export
    dt_str_now = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    file_name = 'analyzed_mails_' + branch + '_' + dt_str_now+ '.xlsx'
    df_analyzed_mails.to_excel(file_name)