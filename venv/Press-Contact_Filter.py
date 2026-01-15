import pandas as pd
import os
from datetime import datetime

folder = r'C:\Users\andre\OneDrive\Desktop\Marketing\ToDo/'
file = 'Verteiler Pressekontakte WMA' + '.xlsx'
########################################################################################################################
# Vorbereitung:
# Alle roten Mails herauslöschen (da Farben nicht beachtet werden können)
# Mit Robinsonliste abgleichen
# Nach Ausführung des Codes kann aus der Ausgabedatei "Presse_Markierung" die Spalte "Markierung" in die Originaldatei eingefügt werden.
# Alle Pressekontakte haben einen Wert größer 0. Je höher der Wert ausfällt, desto eher eignen sie sich für die Kontaktaufnahme.

if __name__ == '__main__':
    os.chdir(folder)
    source_file = pd.read_excel(file)

    col_names = list(source_file.columns)
    keywords =  ['resse', 'press@', 'Medienk', 'edaktion', 'edakteur', 'PR', 'ommunikation', 'ommunication', 'elation', 'Öffentlich',
                 'Media M', 'mediaservice', 'Medien', 'Press', 'precher', 'Werbung', 'Public Affair', 'medien', 'pr@', 'pressoffice']
    exclude = ['promotion']
    new_list = []
    for id, row in source_file.iterrows():
        marked = 0
        last_contact, notes = '', ''
        if 'Bemerkungen' in col_names:
            notes = str(row['Bemerkungen']).lower()
        if 'letzter Kontakt' in col_names:
            last_contact = str(row['letzter Kontakt']).lower()
        company = str(row['Firma']).strip()
        position = str(row['Position']).strip()
        pl = position.lower()
        mail = str(row['richtige eMail']).strip()
        if (any(k in position or k in mail for k in keywords) or 'presse' in notes) and not any(e in position for e in exclude) and not 'kein Presse' in notes:
            marked += 1
        if 'presse' in pl or 'presse' in mail:
            marked += 3
        if 'kommunikation' in pl or 'kommunikation' in mail.lower():
            marked += 1
        if 'ansprech' in pl or 'AP' in position:
            marked += 1
        if any(e in pl for e in ['stv.', 'assisten']) and marked >= 2:
            marked -= 1
        if 'elternzeit' in notes and (str(datetime.now().year) in last_contact or str(datetime.now().year-1) in last_contact):
            marked -= 5
        if 'kein interesse' in notes:
            marked -= 3

        new_list.append([company, marked])
        last_comp = company

    df_marked = pd.DataFrame(new_list,columns=['Firma','Markierung'])
    df_marked.to_excel('Presse_Markierung.xlsx')
    print('finished')

#    for line in new_list:
#        print(line)