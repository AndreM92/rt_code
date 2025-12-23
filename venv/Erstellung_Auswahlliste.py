import pandas as pd
import os
import re
from datetime import datetime

path = r"C:\Users\andre\Documents\Python\rt_code"
source_file = 'Auswahl_Studie Social Media-Performance Beispiel 2025' + '.xlsx'
study = 'Studie Social Media-Performance Beispiel 2025'
########################################################################################################################
# Vorbereitung:
# Alle roten Mails herauslöschen (da Farben nicht beachtet werden können)
# Mit Robinsonliste abgleichen

# Extract text from elements
def extract_text(element):
    if element:
        if not isinstance(element,(str,int,float)):
            element = element.text.strip()
        element = str(element)
        if element == '':
            return element
        elif len(element) >= 1:
            repl_element = element.replace('\u200b','').replace('\xa0', ' ').replace('\\xa0', ' ').replace('\n',' ')
            new_element = re.sub('\s+', ' ', repl_element).strip()
            return new_element
        else:
            return element

def get_variables(row):
    company = extract_text(row['Firma'])
    name = extract_text(row['VN+NN kopiert mit Umlaute'])
    position = extract_text(row['Position'])
    mail = extract_text(row['richtige eMail']).lower()
    contact = str(extract_text(row['letzter Kontakt'])).lower()
    notes = str(extract_text(row['Bemerkungen'])).lower()
    country = str(extract_text(row['Land'])).lower()
    return company, name, position, mail, contact, notes, country

def get_points(company, name, position, mail, contact, notes, country, study):
    points = 0
    # Studienbesteller
    if len(notes) > 50:
        if study.lower() in notes:
            if ('bestell' in notes[:50] or 'bezieht' in notes[:50] or 'kein interess' in notes[:50]) or \
                    'bestell' in contact:
                points -= 99
            else:
                points += 30
        if study[:-2].lower() in notes and ('bestell' in notes[:50] or 'bezieht' in notes[:50]):
            points += 30
    # Studienreihen
    marketing_keys = ['marketing', 'commerce', 'vertrieb', 'seo', 'sea', 'business', 'vertrieb', 'sales', 'key account']
    if 'social media' in study.lower():
        if ('content' in (position.lower() or notes) or 'social' in (position.lower() or notes)):
            points += 5
    else:
        if any(k in position.lower() for k in marketing_keys):
            points += 5
    # Name und Kontakt
    if len(name) > 4:
        points += 1
    if len(notes) > 4:
        points += 1
    if 'bestell' in notes or 'bestell' in contact:
        points += 10
    if 'interess' in notes or 'frag' in notes or 'sprache' in notes or 'anruf' in notes:
        points += 3
    if 'angebot' in notes:
        points += 10
    if 'kein interess' in notes:
        points -= 10
    # Ausländische Firmen ausschließen (?)
    if 'schweiz' in country or 'österreich' in country or 'liechtenstein' in country or 'luxemburg' in country:
        points -= 10
    if 'elternzeit' in notes and (str(datetime.now().year) in contact or str(datetime.now().year-1) in contact):
        points -= 10
    # Eindeutige Pressekontakte ausschließen
    keywords = ['resse', 'press@', 'Medienk', 'edaktion', 'edakteur', 'PR', 'ommunikation', 'ommunication', 'elation',
                'Öffentlich',
                'Media M', 'mediaservice', 'Medien', 'Press', 'precher', 'Werbung', 'Public Affair', 'medien']
    keywords_s = ['resse', 'press@', 'pr@']
#    if (any((k in position or k in mail) for k in keywords) or 'presse' in notes):
    if any(k in mail for k in keywords_s) or 'presse' in notes or 'presse' in position.lower():
        points -= 10
    if not ('marketing' in position.lower() or 'marketing' in mail):
        points += 1
    #Abwesend/ nicht verfügbar?
    if mail[1] == '(':
        points -= 20
    return points


if __name__ == '__main__':
    os.chdir(path)
    contacts_file = pd.read_excel(source_file)
    col_names = list(contacts_file.columns)

    ap_dict = {}

    for id, row in contacts_file.iterrows():
        company, name, position, mail, contact, notes, country = get_variables(row)
        if len(str(mail)) <= 10:
            continue
        points = get_points(company, name, position, mail, contact, notes, country, study)
        new_row = [points] + [v for v in row]
        if company in ap_dict:
            ap_dict[company].append(new_row)
        else:
            ap_dict[company] = [new_row]
    # Vary the positions
    for company, entries in ap_dict.items():
        position_list = []
        for ID, e in enumerate(entries):
            points = e[0]
            position = extract_text(e[13]).lower()
            if len(position) > 4 and any(p in position for p in position_list):
                points -= 10
            position_list.append(position)
            ap_dict[company][ID][0] = points
    # Sort each company's list by points descending
    for company in ap_dict:
        ap_dict[company].sort(key=lambda x: x[0], reverse=True)
    # Remove companies if their lowest score (last entry) is <= -50 (Besteller der jeweiligen Studie)
    ap_dict_f ={}
    for company, entries in ap_dict.items():
        for e in entries:
            if e[0] <= -50:
                continue
            if company in ap_dict_f:
                ap_dict_f[company].append(e)
            else:
                ap_dict_f[company] = [e]
    # Trim each company's list to ten or less
    ap_dict_t = {}
    for company in list(ap_dict_f):
        ap_dict_t[company] = ap_dict_f[company][:10]
    # Transform the entries to a list format
    cleaned_rows = []
    for company, entries in ap_dict_t.items():
        for entry in entries:
            cleaned_rows.append(entry)
    # Create a DataFrame
    columns = ['Points'] + col_names
    df_result = pd.DataFrame(cleaned_rows, columns=columns)
    # Set the 'Points' column as the new index
    df_result.set_index('Points', inplace=True)

    # Export
    dt_str_now = datetime.now().strftime("%Y-%m-%d")
    file_name = 'Auswahl_Aps_' + study + '_' + dt_str_now + '.xlsx'
    df_result.to_excel(file_name)
    print('Done')