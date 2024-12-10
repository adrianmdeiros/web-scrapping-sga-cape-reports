import os
import pandas
import locale
import questionary
from datetime import datetime
from dateutil.relativedelta import relativedelta
from time import sleep
from dotenv import load_dotenv
from splinter import Browser
from tqdm import tqdm

locale.setlocale(locale.LC_TIME,  'pt_BR.utf8')

load_dotenv(override=True)
username = os.getenv('USERNAME')
password = os.getenv('PASSWORD')
username_para = os.getenv('USERNAME_PARA')

actual_date = datetime.now()
year = actual_date.year
prev_month = actual_date - relativedelta(months=1)
last_day_prev_month = (prev_month.replace(day=1) + relativedelta(months=1) - relativedelta(days=1)).day
prev_month_number = prev_month.month
date_object = datetime.strptime(f'01/{prev_month_number}/{year}', '%d/%m/%Y')
month_name = date_object.strftime('%B')

BR_STATES = {
        'AC': 'ACRE',
        'AL': 'ALAGOAS',
        'AP': 'AMAPA',
        'AM': 'AMAZONAS',
        'BA': 'BAHIA',
        'CE': 'CEARA',
        'DF': 'DISTRITO FEDERAL',
        'ES': 'ESPÍRITO SANTO',
        'GO/TO': 'GOIAS',
        'MA': 'MARANHÃO',
        'MT': 'MATO GROSSO',
        'MS': 'MATO GROSSO DO SUL',
        'MG': 'MINAS GERAIS',
        'PA': 'PARÁ',
        'PB': 'PARAÍBA',
        'PR': 'PARANÁ',
        'PE': 'PERNAMBUCO',
        'PI': 'PIAUÍ',
        'RJ': 'RIO DE JANEIRO',
        'RN': 'RIO GRANDE DO NORTE',
        'RS': 'RIO GRANDE DO SUL',
        'RR': 'RORAIMA',
        'SC': 'SANTA CATARINA',
        'SP': 'SAO PAULO',
        'SE': 'SERGIPE'
}


def set_browser(browser):
    return Browser(browser, fullscreen=True)

def click_nav_menu(browser): 
    browser.find_by_css('.nav-link').click()

def open_change_unit_dialog(browser):
    browser.links.find_by_partial_text('Trocar unidade').click()

def get_unit_selector(browser):
    return browser.find_by_id('unidade')

def get_units(unit_selector):
    return unit_selector.find_by_tag('option')

def select_unit(unit_selector, value):
    unit_selector.find_by_value(value).click()

def submit_unit_selection(browser):
    browser.find_by_text('Enviar').click()

def close_change_unit_dialog(browser):
    browser.find_by_id('dialog-unidade').click()

def open_reports_section(browser):
    browser.links.find_by_partial_text('Relatórios').click()

def click_reports_tab(browser):
    browser.links.find_by_partial_href('#tab-relatorios').click()

def get_report_selector(browser):
    return browser.find_by_id('report')

def select_report(report_selector, text):
    report_selector.find_by_text(text).click()

def visit(browser, url):
    browser.visit(url)

def get_unit_state(browser):
    return browser.find_by_tag('h2').first.value

def get_unit_uf(unit_state):
    return unit_state[-5:] if 'GO/TO' in unit_state else unit_state[-2:]

def get_report_table(browser):
    return browser.find_by_tag('table').first

def get_report_table_rows(report_table):
    return report_table.find_by_tag('tr')

def get_data_cells(row):
    return row.find_by_tag('td')

def back(browser):
    browser.back()

def login(browser, username, password):
    browser.fill('username', username)
    browser.fill('password', password)
    browser.find_by_text('Entrar').click()

def click_perfil_menu(browser):
    browser.find_by_xpath('/html/body/header/nav/div/ul[2]/li[3]').click()

def logout(browser):
    browser.links.find_by_partial_text('Sair').click()

def reset_unit_selection_scraping(browser):
    click_nav_menu(browser)
    open_change_unit_dialog(browser)

    sleep(2)

    cape_unit_selector = get_unit_selector(browser)
    cape_units = get_units(cape_unit_selector)
    select_unit(cape_unit_selector, '')
    
    sleep(2)

    close_change_unit_dialog(browser)

    return cape_units

def change_cape_unit_scraping(browser, index):
    sleep(3)

    click_nav_menu(browser)
    
    sleep(2)

    open_change_unit_dialog(browser)

    sleep(2)

    cape_unit_selector = get_unit_selector(browser)
    cape_units = get_units(cape_unit_selector)
    select_unit(cape_unit_selector, cape_units[index].value)
    submit_unit_selection(browser)

def select_report_by_name(browser, name):
    click_nav_menu(browser)
    open_reports_section(browser)
    click_reports_tab(browser)
    
    report_selector = get_report_selector(browser)
    select_report(report_selector, name)

def generate_report(browser, name, i):
    change_cape_unit_scraping(browser, i)

    sleep(3)

    select_report_by_name(browser, name)
    
    visit(browser, f'https://sga.economia.gov.br/novosga.reports/report?report=2&startDate=01%2F{prev_month_number}%2F{year}&endDate={last_day_prev_month}%2F{prev_month_number}%2F{year}')
    
    sleep(3)
    
    state_cape = get_unit_state(browser)
    uf = get_unit_uf(state_cape)

    return uf, state_cape

def get_executed_services_reports(browser, cape_units):
    for i in tqdm(range(1, len(cape_units)), desc="Progresso total", ncols=90, colour='green'):
        uf, state_cape = generate_report(browser, 'Serviços Executados', i)
        
        script = """
            const rows = Array.from(document.querySelectorAll('table tr'));
            return rows
                .map(row => Array.from(row.querySelectorAll('td'))
                    .map(cell => cell.innerText))
                    .filter(cells => cells.every(text => !text.includes('Total'))
                );
        """
        data_cells = browser.execute_script(script)

        print(f'⏳ Lendo tabela/relatório de serviços executados da {state_cape}...')

        for cells in data_cells:
            if cells:
                row_data =  [BR_STATES[uf], f'01/{prev_month_number}/{year}'] + [int(cell) if cell.isdigit() else cell for cell in cells]
                print(f'✅ Linha lida: {row_data}')
                executed_services_rows.append(row_data)

        print(f'✅ Leitura finalizada. Partindo para a próxima unidade...')

        back(browser)

def get_parah_executed_services_report(browser):
    click_perfil_menu(browser)
    logout(browser)

    sleep(3)

    login(browser, username_para, password)

    sleep(3)

    cape_units = reset_unit_selection_scraping(browser)
    get_executed_services_reports(browser, cape_units)

def executed_services_scraping(browser):
    visit(browser,'https://sga.economia.gov.br/')

    sleep(3)

    login(browser, username, password)

    sleep(3)

    cape_units = reset_unit_selection_scraping(browser)
    get_executed_services_reports(browser, cape_units)

    sleep(3)
    
    print('👽 Iniciando o caso do PARÁ...')
    get_parah_executed_services_report(browser)

def get_finished_services_reports(browser, cape_units):
    pass

def get_parah_finished_services_report(browser, cape_units):
    pass

def finished_services_scraping(browser):
    pass

def format_excel(styler):
    styler.set_properties(**{"font-size": "10pt", "font-family": "Segoe UI"})
    return styler

if __name__ == '__main__':
    executed_services_rows = [['Estado', 'mês/ano', 'Serviço', 'Quantidade']]
    completed_services_rows = [['Senha', 'Data', 'Chamada', 'Início', 'Fim', 'Duração', 'Permanência', 'Serviço Triado', 'Atendente', 'Estado']]
    
    selected = questionary.select(
        "Qual relatório você quer rodar?",
        choices=["Serviços Executados", "Atendimentos Conluídos"]
    ).ask()
    
    browser = set_browser('chrome')
    
    if selected == "Serviços Executados":
        executed_services_scraping(browser)
        executed_services_table = pandas.DataFrame(executed_services_rows)

        print(f'⏳ Gerando arquivo Excel...')

        executed_services_table.style.pipe(format_excel).to_excel(f'SERVICOS EXECUTADOS CAPES {month_name.upper()} {year}.xlsx', index=False, header=False)
        
        print(f'✅ Arquivo Excel criado.')
    
    elif selected == "Atendimentos Conluídos":
        finished_services_scraping(browser)
        completed_services_table = pandas.DataFrame(completed_services_rows)

        print(f'⏳ Gerando arquivo Excel...')

        completed_services_table.style.pipe(format_excel).to_excel(f'ATENDIMENTOS CONCLUIDOS CAPES {month_name.upper()} {year}.xlsx', index=False, header=False)

        print(f'✅ Arquivo Excel criado.')

    browser.quit()

        
    
    
    

    