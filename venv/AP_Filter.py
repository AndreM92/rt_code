
import pandas as pd
import os
import re
from datetime import datetime

path = r"C:\Users\andre\OneDrive\Desktop\Marketing\ToDo/"

# Kontaktliste verkleinern
all_contacts_file = 'Kontakte_Versicherungen'
p_contacts = path + '/' + all_contacts_file + ".xlsx"
p_distinct_contacts = p_contacts
########################################################################################################################

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

def get_source_file_vars(rown, col_list):
    brand, full_name, full_name2, ad_volume = ['' for _ in range(4)]
    name_variations = ['Firma', 'Unternehmen', 'Shops']
    for n in name_variations:
        if n in col_list:
            full_name = extract_text(rown[n])
    if 'Firma2' in col_list:
        full_name2 = extract_text(rown['Firma'])
    if 'Marke' in col_list:
        brand = extract_text(rown['Marke'])
        if not brand or 'nan' in str(brand):
            brand = ''
        else:
            if '/' in brand:
                bn = brand.split('/')[0]
                if len(bn) >= 4:
                    brand = bn
    # eV Studien
    if 'eV' in new_contacts_file:
        if '.' in full_name:
            brand = full_name.split('.')[0].strip()
    for e in col_list:
        if 'ZR 5' in e or 'Werbeausgaben' in e or "Werbevolumen" in e or "Σ" in e:
            ad_volume = extract_text(rown[e])
            if ad_volume and ad_volume[0].isdigit():
                ad_volume = round(float(ad_volume))
    return brand, full_name, full_name2, ad_volume

def create_keywords(full_name, brand):
    if full_name:
        c_name = full_name
    elif brand:
        c_name = brand
    else:
        return '',[]
    adds = ['AB', 'a.s.', 'S.A.', ' OG', 'AG', ' SE', 'GmbH & Co. KG', 'GmbH', 'B.V.', 'KG', 'LLC', 'NV', 'N.V.',
            '& Co.', 'S.L.U.', '(', ')', '.de', '.com', '.at', 'oHG', 'Ltd.', 'Limited']
    for a in adds:
        c_name = c_name.replace(a,'')
    c_name = c_name.strip()
    name_list = [e.strip() for e in c_name.split() if (len(e.strip()) >= 4 and e.strip().lower() != 'nan')]
    name_list_special = [e.strip() for e in c_name.split('/') if (len(e.strip()) >= 4 and e.strip().lower() != 'nan')]
    if brand and len(brand) >= 3:
        name_list.append(brand.strip())
        name_list.append(brand.lower().strip())
    name_var = c_name.replace('-',' ').replace('_',' ').replace('.',' ')
    name_var_list = [n.strip() for n in name_var.split(' ') if len(n) >= 4]
    name_list += name_var_list
    name_list = list(set(name_list))
    #Special filter
#    f_list = ['bank', 'banken']
#    name_list_filter = [e for e in name_list if e.lower() not in f_list]
    return c_name, name_list


if __name__ == '__main__':
    os.chdir(path)
    contacts_df = pd.read_excel(p_contacts)
    distinct_contacts = []
    dist_firms = []
    for id, row in contacts_df.iterrows():
        comp = extract_text(row['Firma'])
        if comp not in dist_firms:
            dist_firms.append(comp)
            distinct_contacts.append(row)

    df_dist_cont = pd.DataFrame(distinct_contacts,columns=contacts_df.columns)
    p_distinct_contacts = all_contacts_file + '_gekürzt' + ".xlsx"
    df_dist_cont.to_excel(p_distinct_contacts)

########################################################################################################################
# Zu überprüfende Firmen
new_contacts_file = 'Liste Firmen_WMA ÖPNV'
p_newc = r"C:\Users\andre\OneDrive\Desktop\Marketing\ToDo/" + new_contacts_file + ".xlsx"

# Volle Kontaktliste
if __name__ == '__main__':
    os.chdir(path)
    p_distinct_contacts = all_contacts_file + '_gekürzt' + ".xlsx"
    df_dist_cont = pd.read_excel(p_distinct_contacts)
    new_companies = pd.read_excel(p_newc)
    col_list = list(new_companies.columns)

    checked_companies = []
    for idx, rown in new_companies.iterrows():
        brand, full_name, full_name2, ad_volume = get_source_file_vars(rown, col_list)
        if not brand and not full_name:
            continue

        full_name_l = full_name.lower()
        full_name_l2 = full_name2.lower()
        bl = brand.lower()

        found1, found2, found3, found4 = ['' for _ in range(4)]
        for id, row in df_dist_cont.iterrows():
            rownumber = id + 2
#            if rownumber == 400:
#                break
            c_company = extract_text(row['Firma'])
            branch = ''
            for c in df_dist_cont.columns:
                if 'Branche' in c:
                    branch = extract_text(row[c])
                    break

            found_string = str(rownumber) + ': ' + c_company + '_' + branch

            if not c_company or len(str(c_company)) <= 4:
                continue
            homepage = extract_text(row['Hinweis-Homepage']).lower()
            hs = homepage.split('.')[0].strip()
            if brand and len(brand) >= 3 and full_name and len(full_name) >= 4:
                c_company_s, name_list = create_keywords(c_company, None)
                c_company_sl = c_company_s.lower()
                if c_company_sl in full_name_l or c_company_sl in bl or brand in c_company_s or \
                        c_company_sl in full_name_l2 or bl in full_name_l2 or (len(brand) > 4 and bl in c_company_sl)\
                        or full_name in homepage:
                    if found1:
                        found1 = found1 + ', ' + found_string
                    else:
                        found1 = found_string
                if c_company == found1 or len(found1) > 200 or len(found2) >= 1000:
                    continue
                if (bl in c_company_sl or c_company_sl in bl or bl in homepage) and c_company not in found1:
                    if found2 and c_company not in found1:
                        found2 = found2 + ', ' + found_string
                    else:
                        found2 = found_string
                if len(hs) >= 4:
                    if hs in bl:
                        if found2:
                            found2 = found2 + ', ' + found_string
                        else:
                            found2 = found_string

            elif brand and len(brand) >= 3:
                if len(c_company) - len(brand) <= 20:
                    if brand in c_company or c_company in brand or (len(brand) > 4 and bl in c_company_sl):
                        if found1:
                            found1 = found1 + ', ' + found_string
                        else:
                            found1 = found_string
                if c_company in found1 or len(found2) >= 1000 or len(c_company) - len(brand) >= 30:
                    continue
                if (bl in c_company_sl or c_company_sl in bl or bl in homepage) and c_company not in found1:
                    if found2 and c_company not in found1:
                        found2 = found2 + ', ' + found_string
                    else:
                        found2 = found_string
                if len(hs) >= 4:
                    if hs in bl:
                        if found2:
                            found2 = found2 + ', ' + found_string
                        else:
                            found2 = found_string

            # Only company names
            elif full_name and len(full_name) >= 3:
                new_company_s, new_name_list = create_keywords(full_name, None)
                new_company_sl = new_company_s.lower()
                c_company_s, name_list = create_keywords(c_company, None)
                c_company_sl = c_company_s.lower()
                if c_company_sl == new_company_sl:
                    if found1:
                        found1 = found1 + ', ' + found_string
                    else:
                        found1 = found_string
                elif new_company_sl in c_company_sl or c_company_sl in new_company_sl or (new_company_sl in hs and len(hs) >= 4):
                    if found2 and c_company not in found1:
                        found2 = found2 + ', ' + found_string
                    else:
                        found2 = found_string

#        if not found1 and not found2:
        c_name, name_list = create_keywords(full_name, brand)
        cl_name = c_name.lower()
        if len(name_list) == 0:
            continue
        for id, row in df_dist_cont.iterrows():
            rownumber = id + 2
            c_company = extract_text(row['Firma'])
            hl = extract_text(row['Hinweis-Homepage']).lower()
            hs = hl.split('.')[0].strip()
            if len(hs) >= 4:
                c_name_2, name_list_2 = create_keywords(c_company, hs)
            elif len(brand) >= 4:
                c_name_2, name_list_2 = create_keywords(c_company, brand)
            else:
                c_name_2, name_list_2 = create_keywords(c_company, None)
            cl_name_2 = c_name_2.lower()
            if any(n in cl_name_2 for n in name_list):
                found3_string = str(rownumber) + ': ' + c_company + ', '
                found3 += found3_string
            if c_company not in found3 and any(n in cl_name for n in name_list_2):
                found4_string = str(rownumber) + ': ' + c_company + ', '
                found4 += found4_string
        output_row = [idx, full_name, brand, ad_volume, found1, found2, found3, found4]
        checked_companies.append(output_row)
        print(output_row)
#        if idx >= 4:
#            break

    df_checked_comp = pd.DataFrame(checked_companies,columns=['ID','Firma','Marke','Werbevolumen/Punkte','Suche1','Suche2','Suche3','Suche4'])

    # Export
    dt_str_now = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    file_name = 'APs_' + new_contacts_file + '_' + dt_str_now + '.xlsx'
    df_checked_comp.to_excel(file_name)
    print('Done')