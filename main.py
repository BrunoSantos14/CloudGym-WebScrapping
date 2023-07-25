from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv
import time
import os
import re


load_dotenv(override=True)

EMAIL = os.getenv("EMAIL")
SENHA = os.getenv("SENHA")


class CloudGym:
    """
    Web Scrapping dos dados da academia da Controllab. Inicializa com o período selecionado pelo usuário:
    periodo: default -> 'ontem' - {'hoje', 'ontem', 'semana_atual', 'semana_passada', 'mes_atual' 'mes_passado'}.

    Obs.: Não existe tratamento de ano para periodos com semana (semana_atual, semana_passada e proxima_semana)
    """
    
    def __init__(self, periodo='ontem'):
        self.periodo = periodo
        self.caminho = '//Pastor/analytics/Estudos/Academia'
        # options = webdriver.ChromeOptions()
        # options.add_argument("--headless=new")
        # servico = Service(ChromeDriverManager().install()) 
        self.driver = webdriver.Chrome() #service=servico) #, options=options)
        self.driver.maximize_window()


        self.__login()
        self.__navegar_ocupacao()
        self.__filtrar()
        self.dados = self.__raspagem()
        self.agrupar_tabelas
    

    def __login(self):
        """Faz login no site da academia.""" 
        url = 'https://app.cloudgym.io/'
        self.driver.get(url)

        self.driver.find_element(By.CLASS_NAME, "form-control")\
            .send_keys(EMAIL)

        self.driver.find_element(By.ID, "passwd")\
            .send_keys(SENHA)
        
        self.driver.find_element(By.CLASS_NAME, 'btn')\
            .click()
    

    def __navegar_ocupacao(self):
        """Navega até a página de ocupação (onde tem chek-in feito pelo professor da academia)"""
        espera = WebDriverWait(self.driver, 10)
        espera.until(EC.presence_of_element_located((By.CLASS_NAME, 'fitness'))).click()

        time.sleep(0.3)
        self.driver.find_element(By.ID, 'liattendance').click()
    
    
    def __filtrar(self):
        """Filtra o período escolhido."""

        # Acionando lista de períodos
        self.driver.find_element(By.ID, 'classAttendanceReportRange').click()

        # Pegando xpath do periodo escolhido
        dic = {
            'hoje': 1,
            'ontem': 2,
            'amanha': 3,
            'semana_atual': 4,
            'semana_passada': 5,
            'proxima_semana': 6,
            'mes_atual': 7,
            'mes_passado': 8,
        }
        num = dic.get(self.periodo)
        xpath = f'/html/body/div[5]/div[1]/ul/li[{num}]'

        self.driver.find_element(By.XPATH, xpath).click()
        time.sleep(1)


    def __encontrar_ano(self):
        """Encontra o ano da data com base na data atual."""
        if self.periodo == 'ontem':
            return str((datetime.now() - relativedelta(days=1)).year)

        elif self.periodo == 'hoje' or self.periodo == 'mes_atual':
            return str((datetime.now()).year)

        elif self.periodo == 'amanha':
            return str((datetime.now() + relativedelta(days=1)).year)

        elif self.periodo == 'mes_passado':
            return str((datetime.now() - relativedelta(months=1)).year)

        elif self.periodo == 'semana_passada':
            return str((datetime.now()).year)


    def __raspagem(self):
        """Encontra todos os dias, horas, professores, alunos e seu check-in de todo o período selecionado.
        Retorna uma lista com tuplas para auxiliar na montagem de DataFrame."""

        tabela = self.driver.find_elements(By.CLASS_NAME, 'table-scrollable')[1]

        linhas = tabela.text.split('Participantes')
        linhas = [i.split('\n') for i in linhas]
        linhas = [''.join(i) for i in linhas]

        datas = [i[i.find('/')-2 : i.find('/')+3] for i in linhas]
        horas = [i[i.find(':')-2 : i.find(':')+3] for i in linhas]
        lista = [i[i.find(':00 ')+4 : ] for i in linhas]

        professores = []
        for item in linhas:
            match = re.search(r'([A-Za-z]+)\d', item)
            if match:
                professores.append(match.group(1))

        # Filtrando datas necessárias
        datas = [data + f'/{self.__encontrar_ano()}' for data in datas if len(data) > 0]
        datas = [datetime.strptime(data, '%d/%m/%Y').date() for data in datas]
        datas = [data for data in datas if data < datetime.now().date()]
        datas = [data.strftime('%d/%m/%Y') for data in datas]

        # Achando os botões com a lista dos participantes
        btn_part = self.driver.find_elements(By.CLASS_NAME, 'fa-list')

        # Padronizando tamanhos das listas
        horas = horas[:len(datas)]
        professores = professores[:len(datas)]
        btn_part = btn_part[:len(datas)]

        lista = []
        for i, popup in enumerate(btn_part):
            # Abrindo o popup
            popup.click()
            time.sleep(1)
            
            # Guardando o nome e status do participante
            nomes = [i.text for i in self.driver.find_elements(By.CLASS_NAME, 'gradeX')]
            status = [i.get_attribute('style') for i in self.driver.find_elements(By.CLASS_NAME, 'text-center') 
                      if i.get_attribute('style') != '']
            status = ['Sim' if i == 'color: mediumslateblue;' else 'Não' for i in status]
            
            lista.append(
                (datas[i], horas[i], professores[i], nomes, status)
                )
            
            # Fechando popup
            self.driver.find_element(By.CLASS_NAME, 'close').click()
            time.sleep(1)
        return lista


    def __transformar_tabela(self):
        """Transforma a lista obtida na raspagem em um DataFrame."""

        colunas = ['data', 'hora', 'professor', 'aluno', 'check_in']
        return pd.DataFrame(self.dados, columns=colunas).explode(['aluno', 'check_in'])

    @property
    def agrupar_tabelas(self):
        """Agrupa arquivo salvo no diretório com os novos atualizados."""
        # Pegando dados novos
        df = self.__transformar_tabela()

        # Lendo o arquivo antigo
        academia = pd.read_excel(self.caminho + '/academia.xlsx')

        # Retirando datas encontradas na raspagem (se precisar)
        academia = academia[~academia['data'].isin(df['data'])]

        # Agrupando os dataframes 
        agrupado = pd.concat([academia, df])
        agrupado.to_excel(self.caminho + '/academia.xlsx', index=False)

        self.driver.quit()



def main():
    """
    Segunda -> rodar pegando semana passada.
    Terça a sexta -> rodar pegando ontem.
    """
    if datetime.now().weekday() == 0:
        CloudGym('semana_passada')

    elif datetime.now().weekday() in [1, 2, 3, 4]:
        CloudGym('ontem')


if __name__ == '__main__':
    main()
