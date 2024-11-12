import os
import pandas
import locale
from datetime import datetime
from dateutil.relativedelta import relativedelta
from time import sleep
from dotenv import load_dotenv
from splinter import Browser

locale.setlocale(locale.LC_TIME,  'pt_BR.utf8')
actual_date = datetime.now()
year = actual_date.year
prev_month = actual_date - relativedelta(months=1)
prev_month_number = prev_month.month
last_day_prev_month = (prev_month.replace(day=1) + relativedelta(months=1) - relativedelta(days=1)).day

states = {
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

completed_services_rows = [
    ['Senha', 'Data', 'Chamada', 'Início', 'Fim', 'Duração', 'Permanência', 'Serviço Triado', 'Atendente']
]

browser = Browser('chrome', fullscreen=True)

def get_reports():
    browser.find_by_css('.nav-link').click()
    browser.links.find_by_partial_text('Trocar unidade').click()

    sleep(2)

    cape_units_selector = browser.find_by_id('unidade')
    cape_units = cape_units_selector.find_by_tag('option')
    cape_units_selector.find_by_value('').click()

    sleep(2)

    browser.find_by_id('dialog-unidade').click()

    for i in range(1, len(cape_units)):
        sleep(3)

        browser.find_by_css('.nav-link').click()

        sleep(2)

        browser.links.find_by_partial_text('Trocar unidade').click()

        sleep(2)
        
        cape_units_selector = browser.find_by_id('unidade')
        cape_units = cape_units_selector.find_by_tag('option')
        cape_units_selector.find_by_value(cape_units[i].value).click()
        browser.find_by_text('Enviar').click()

        sleep(2)

        browser.find_by_css('.nav-link').click()
        browser.links.find_by_partial_text('Relatórios').click()
        browser.links.find_by_partial_href('#tab-relatorios').click()
        report_selector = browser.find_by_id('report')
        report_selector.find_by_text('Atendimentos concluídos').click()
        print('⏳ Gerando relatório. Por favor aguarde...')
        browser.visit(f'https://sga.economia.gov.br/novosga.reports/report?report=3&startDate=01%2F{prev_month_number}%2F{year}&endDate={last_day_prev_month}%2F{prev_month_number}%2F{year}')
        

        sleep(3)
    
        state_cape = browser.find_by_tag('h2').first.value

        uf = ''
        if state_cape == 'CAPE - GO/TO':
            uf = state_cape[-5:]
        else:
            uf = state_cape[-2:]

        table = browser.find_by_tag('table').first
        
        print(f'⏳ Lendo tabela/relatório de atendimentos concluídos da CAPE-{uf}. Esse processo pode demorar um pouco...')
        
        for row in table.find_by_tag('tr'):
            data_cells = row.find_by_tag('td')
            
            if data_cells and len(data_cells) == 9:
                row_data = [cell.text for cell in data_cells if 'Total' not in cell.text]
                completed_services_rows.append(row_data)
                print(row_data)
                    

        print(f'✅ Leitura finalizada. Partindo para a próxima unidade...')

        browser.back()

browser.visit('https://sga.economia.gov.br/')

load_dotenv(override=True)
username = os.getenv('USERNAME')
password = os.getenv('PASSWORD')

browser.fill('username', username)
browser.fill('password', password)
browser.find_by_text('Entrar').click()

sleep(3)

get_reports()

sleep(3)

print(f'👽 Iniciando o misterioso caso do PARÁ...')

browser.find_by_xpath('/html/body/header/nav/div/ul[2]/li[3]').click()
browser.links.find_by_partial_text('Sair').click()

sleep(3)

username_para = os.getenv('USERNAME_PARA')
browser.fill('username', username_para)
browser.fill('password', password)
browser.find_by_text('Entrar').click()

sleep(3)

get_reports()

print(f'✅ Todas as tabelas/relatórios lidos com sucesso.')
print(f'⏳ Gerando arquivo Excel...')

completed_services_table = pandas.DataFrame(completed_services_rows)

date_object = datetime.strptime(f'01/{prev_month_number}/{year}', '%d/%m/%Y')
month_name = date_object.strftime('%B')

completed_services_table.to_excel(f'ATENDIMENTOS CONCLUIDOS CAPES {month_name.upper()} {year}.xlsx', index=False, header=False)

print(f'✅ Arquivo Excel criado com sucesso!🚀')

browser.quit()