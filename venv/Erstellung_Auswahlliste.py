import pandas as pd
import os
import re
from datetime import datetime
from openpyxl import load_workbook

path = r"C:\Users\andre\Documents\Python\rt_code"
source_file = 'Auswahl_WMA Kfz-Versicherung 2026_vorläufig' + '.xlsx'
study = 'Werbemarktanalyse Kfz-Versicherung 2026'
positivliste = ['online', 'werbung']
negativliste = ['assisten', 'produkt']

########################################################################################################################
#source_file = 'Auswahl_Studie Social Media-Performance Beispiel 2025' + '.xlsx'
#study = 'Studie Social Media-Performance Beispiel 2025'

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
            new_element = re.sub(r'\s+', ' ', repl_element).strip()
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

def get_points(name, position, mail, contact, notes, country, study, font_color, positivliste, negativliste):
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
    # Ausländische Firmen abwerten oder ausschließen
    if len(country) > 4 and not 'deutschland' in country:
        if ('schweiz' in country or 'österreich' in country or 'liechtenstein' in country or 'switzerland' in country
                or 'austria' in country):
            points -= 5
        else:
            points -= 15
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
    # Abwesend/ nicht verfügbar?
    if mail[1] == '(':
        points -= 20
    # Schriftfarbe
    if 'FF4F81BD' in font_color:
        points += 20
    if 'FF9BBB59' in font_color:
        points += 10
    if 'FFFF0000' in font_color:
        points -= 999
    # Eigene Positiv- und Negativliste
    if any(e in position.lower() for e in positivliste):
        points += 10
    if any(e in position.lower() for e in negativliste):
        points -= 10
    return points
########################################################################################################################

if __name__ == '__main__':
    os.chdir(path)
    df_contacts_file = pd.read_excel(source_file, engine="openpyxl")
    col_names = list(df_contacts_file.columns)
    # Excel mit openpyxl laden (für Formatierung)
    wb = load_workbook(source_file, data_only=False)
    ws = wb.active
    ap_dict = {}
    company_appendix = ['GmbH & Co. KG', 'gmbh', 'mbh', 'inc', 'limited', 'ltd', 'llc', 'co.', 'lda', 'a.s.', 'SE'
                        'S.A.', ' OG', ' AG', ' SE', 'GmbH', 'B.V.', 'KG', 'LLC', 'NV', 'N.V.', '& Co.', 'S.L.U.',
                        '(', ')', '.de', '.com', '.at', 'oHG', 'Ltd.', 'Limited', 'eG', 'P.S.K.', 'S.p.A.']
    press_keywords = ['resse', 'press@', 'Medienk', 'edaktion', 'edakteur', 'PR', 'ommunikation', 'ommunication',
                      'elation', 'Öffentlich', 'Media M', 'mediaservice', 'Medien', 'Press', 'precher', 'Werbung',
                      'Public Affair', 'medien']
    pks = ['resse', 'press@', 'pr@']
    for id, row in df_contacts_file.iterrows():
        company, name, position, mail, contact, notes, country = get_variables(row)
        if len(str(mail)) <= 10:
            continue
        # Index der Spalte "richtige eMail" finden
        email_col = next(i for i, c in enumerate(ws[1], start=1)
            if c.value and str(c.value).strip().lower() == "richtige email")
        # entsprechende Zelle und Schriftfarbe holen
        cell = ws.cell(row=id + 2, column=email_col)  # +1 Header, +1 von 0- auf 1-basiert
        font_color = str(getattr(getattr(cell.font, "color", None), "rgb", None) or "NO_COLOR")
        points = get_points(name, position, mail, contact, notes, country, study, font_color, positivliste, negativliste)
        new_row = [points] + [v for v in row]
        company_s = company
        for a in company_appendix:
            company_s = company_s.replace(a,'').strip()
        # Kürzung des Firmen-Keywords auf den ersten Namen:
        if company_s.find(' ') >= 3:
            company_s = company_s.split()[0]
        if company_s in ap_dict:
            ap_dict[company_s].append(new_row)
        else:
            ap_dict[company_s] = [new_row]
    # Vary the positions and email structures
    for company, entries in ap_dict.items():
        position_list = []
        marketing_h = 0
        marketing_hs = 0
        press_h = 0
        press_sh = 0
        for ID, e in enumerate(entries):
            mail = str(e[17])
            points = e[0]
            position = extract_text(e[13])
            # Mail variations
            if ID < 1:
                if mail.find('.') > 2:
                    pointpos_short = False
                else:
                    pointpos_short = True
            if ID >= 1:
                if mail.find('.') > 2:
                    if pointpos_short == True:
                        points += 10
                        pointpos_short = False
                else:
                    if pointpos_short == False:
                        points += 10
                        pointpos_short = True
            if not position or len(position) < 4:
                continue
            if any(k in position for k in press_keywords) or any(k.lower() in mail for k in press_keywords) or \
                    any('presse' in str(t).lower() for t in e):
                if any(e.lower() in position for e in pks) or any(e.lower() in mail for e in pks):
                    if press_sh >= 1:
                        points -= 10
                    press_sh += 1
                elif press_h >= 3 and any(k in position for k in pks) and press_h >= 1:
                    points -= 10
                press_h += 1
            if any(p in position for p in position_list):
                points -= 5
            if 'marketing' in position.lower():
                if marketing_h >= 3:
                    points -= 10
                marketing_h += 1
                if position.lower() == 'marketing':
                    if marketing_hs >= 1:
                        points -= 10
                    marketing_hs += 1
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
    for company, entries in ap_dict_f.items():
        n = len(entries)
        if n >= 10:
            limit = 10
        elif n >= 7:
            limit = 7
        elif n >= 5:
            limit = 5
        else:
            limit = 3
        ap_dict_t[company] = entries[:limit]
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