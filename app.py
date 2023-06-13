import streamlit as st
import requests
import json
import io
import utils
import locale
from docx import Document
from docx.shared import Pt
import datetime

utils.set_page_configuration(initial_sidebar_state="collapsed")
utils.set_sidebar()

# Pappers setup
url = "https://api.pappers.fr/v2/entreprise"
api_token = utils.open_file('pappers_api_key.txt')


title = 'Company sheet generator'
subtitle = "Automatically create the Company Sheet from a SIRET"
mission = "Connected to Pappers API, this app automatically generates the Company Sheet from a SIRET you provide. The only thing you have to do is to enter the SIRET of the company you are interested in."
st.markdown(f"<h1 style='text-align: center;font-size:50px;'>{title}</h1>", unsafe_allow_html=True)
st.markdown(f"<h3 style='text-align: center;'><i>{subtitle}</i></h3>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: justify;'>{mission}</p>", unsafe_allow_html=True)

st.text('')
siret = st.text_input('SIRET:', placeholder='Ex: 34300986600611', label_visibility="visible")
st.text('')
st.text('')

cols = st.columns(3)
run = cols[1].button('Generate Company Sheet', key=None, help=None, on_click=None, args=None, kwargs=None, disabled=False)

if run:
    st.text('')

    params = {"api_token": api_token, "siret": siret}
    response = requests.get(url, params=params)

    if response.status_code == 200:
        data = response.json()
        # Process the retrieved data as needed
        print(data)
    
        
        # data = json.load(open('data/google.json', 'r', encoding='utf-8'))
        doc = Document('assets/template_fiche_societe.docx')

        # Create distribution of dividends
        three_lastest_exercices = [el['annee'] for el in data['finances']]
        three_lastest_exercices.sort(reverse=True)
        three_lastest_exercices = three_lastest_exercices[:3]

        # Create fiscal year
        locale.setlocale(locale.LC_TIME, "fr_FR")
        end_date = datetime.datetime.strptime(data['date_cloture_exercice'], "%d %B").strftime("%d %B")
        start_date = (datetime.datetime.strptime(data['date_cloture_exercice'], "%d %B") + datetime.timedelta(days=1)).strftime("%d %B")
        fiscal_year = f'{start_date} au {end_date}'

        # Create distribution of dividends
        lastest_exercices = [el['annee'] for el in data['finances']]
        lastest_exercices.sort(reverse=True)
        if len(three_lastest_exercices) >= 3:
            three_lastest_exercices = lastest_exercices[:3]
            distribution_of_dividends = {
                'year_1': str(three_lastest_exercices[0]),
                'net_income_1': str([el['chiffre_affaires'] for el in data['finances'] if el['annee'] == three_lastest_exercices[0]][0]),
                'dividends_1': '', 'distribution_date_1': '',
                'year_2': str(three_lastest_exercices[1]),
                'net_income_2': str([el['chiffre_affaires'] for el in data['finances'] if el['annee'] == three_lastest_exercices[1]][0]),
                'dividends_2': '', 'distribution_date_2': '',
                'year_3': str(three_lastest_exercices[2]),
                'net_income_3': str([el['chiffre_affaires'] for el in data['finances'] if el['annee'] == three_lastest_exercices[2]][0]),
                'dividends_3': '', 'distribution_date_3': ''
            }
        elif len(three_lastest_exercices) == 2:
            two_lastest_exercices = lastest_exercices[:2]
            distribution_of_dividends = {
                'year_1': two_lastest_exercices[0],
                'net_income_1': str([el['chiffre_affaires'] for el in data['finances'] if el['annee'] == two_lastest_exercices[0]][0]),
                'dividends_1': '', 'distribution_date_1': '',
                'year_2': two_lastest_exercices[1],
                'net_income_2': str([el['chiffre_affaires'] for el in data['finances'] if el['annee'] == two_lastest_exercices[1]][0]),
                'dividends_2': '', 'distribution_date_2': '', 'year_3': '', 'net_income_3': '', 'dividends_3': '', 'distribution_date_3': ''
            }
        elif len(three_lastest_exercices) == 1:
            lastest_exercice = lastest_exercices[0]
            distribution_of_dividends = {
                'year_1': lastest_exercice,
                'net_income_1': str([el['chiffre_affaires'] for el in data['finances'] if el['annee'] == lastest_exercice][0]),
                'dividends_1': '', 'distribution_date_1': '', 'year_2': '', 'net_income_2': '', 'dividends_2': '', 
                'distribution_date_2': '', 'year_3': '', 'net_income_3': '', 'dividends_3': '', 'distribution_date_3': ''
            }
        else:
            distribution_of_dividends = {
                'year_1': '', 'net_income_1': '', 'dividends_1': '', 'distribution_date_1': '', 'year_2': '', 'net_income_2': '', 'dividends_2': '', 
                'distribution_date_2': '', 'year_3': '', 'net_income_3': '', 'dividends_3': '', 'distribution_date_3': ''
            }


        replacements = {
            'country': data['siege']['pays'],
            'company_name': data['nom_entreprise'],
            'corporate_form': data['forme_juridique'],
            'registered_office': data['siege']['adresse_ligne_1'] + ', ' + data['siege']['code_postal'] + ', ' + data['siege']['ville'] + ', ' + data['siege']['pays'],
            'share_capital': data['capital_formate'],
            # 'issued_shares': 'Not provided',
            # 'shareholding': 'Not provided',
            'registration_number': data['siren_formate'],
            'rcs_inscription': f"{data['statut_rcs']} (au greffe de {data['greffe']}, le {data['date_immatriculation_rcs']})",
            'company_purpose': data['objet_social'],
            'term': (datetime.datetime.strptime(data['date_immatriculation_rcs'], '%Y-%m-%d') + datetime.timedelta(days=data['duree_personne_morale']*365)).strftime('%Y-%m-%d'),
            'fiscal_year': fiscal_year,
            # 'management_president': [x for x in data['representants'] if x['qualite'] == 'Président'][0]['nom_complet'],
            # 'management_board_of_directors': [x['nom_complet'] for x in data['representants'] if x['qualite'] == 'Gérant'],
            'management': ", ".join([f"{el['nom_complet']} ({el['qualite']})" for el in data['representants'] if 'Commissaire aux comptes' not in el['qualite']]),
            'statutory_auditors_principals': "Principals: " + " ".join([f"{el['nom_complet']} ({el['siren']})" for el in data['representants'] if 'Commissaire aux comptes titulaire' in el['qualite']]),
            'statutory_auditors_alternates': "Alternates: " + " ".join([f"{el['nom_complet']} ({el['siren']})" for el in data['representants'] if el['qualite'] == 'Commissaire aux comptes suppléant']),
            # 'statutory_restrictions_on_transfer_of_shares': 'Not provided',
            # 'powers_restriction_board_of_directors_reserved_matters': 'Not provided',
            # 'other_securities_giving_access_to_the_share_capital': 'Not provided',
            # 'pledge_over_securities': 'Not provided',
            'year1': distribution_of_dividends['year_1'],
            'net_income1': distribution_of_dividends['net_income_1'],
            'dividends1': distribution_of_dividends['dividends_1'],
            'distributiondate1': distribution_of_dividends['distribution_date_1'],
            'year2': distribution_of_dividends['year_2'],
            'net_income2': distribution_of_dividends['net_income_2'],
            'dividends2': distribution_of_dividends['dividends_2'],
            'distributiondate2': distribution_of_dividends['distribution_date_2'],
            'year3': distribution_of_dividends['year_3'],
            'net_income3': distribution_of_dividends['net_income_3'],
            'dividends3': distribution_of_dividends['dividends_3'],
            'distributiondate3': distribution_of_dividends['distribution_date_3'],
            'date_of_updated_articles_of_association': data['derniere_mise_a_jour_sirene']
        }


        for paragraph in doc.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    for run in paragraph.runs:
                        if key in run.text:
                            run.text = run.text.replace(key, value)
                            run.font.name = 'Arial'
                            run.font.size = Pt(8)
                            run.font.bold = True


        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in replacements.items():
                        if key in cell.text:
                            # Replace the field and set font style and size
                            paragraphs = cell.paragraphs
                            for paragraph in paragraphs:
                                run = paragraph.runs
                                for i in range(len(run)):
                                    if key in run[i].text:
                                        run[i].text = run[i].text.replace(key, value)
                                        run[i].font.name = 'Arial'
                                        run[i].font.size = Pt(8)

        # doc.save('results/filled_document.docx')
        bio = io.BytesIO()
        doc.save(bio)

        st.text('')
        st.text('')
        st.markdown(f"<h3 style='text-align: center;'>The Company Sheet is ready (already) 🎁</h3>", unsafe_allow_html=True)
        st.text('')
        cols = st.columns(3)
        cols[1].download_button(
            label="Download Company Sheet",
            data=bio.getvalue(),
            file_name=f"{replacements['company_name'].replace(' ', '_').lower()}_sheet.docx",
            mime="docx",
            key='docx')
    
    elif response.status_code == 400:
        st.write(f"⚠️ The SIRET number {siret} is not valid. Please check the number and try again.")
        st.write(f"Make sure you enter the 14-digit SIRET number, not the 9-digit SIREN number.")
        st.write(f"Also, make sure you enter the number without spaces or dashes.")

    else:
        st.write(f"Error occurred: {response.status_code}")    