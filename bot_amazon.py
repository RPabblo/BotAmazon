from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, NoSuchAttributeException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.chrome.service import Service
import xlsxwriter
from time import sleep
from colorama import Fore, Style
from datetime import datetime
import sys


def line():
    # Função que mostra na tela uma linha. Apenas um detalhe :D
    print('–' * 42)


def readint(n):
    # Função que recebe um input do usuário e retorna o módulo do valor somente se ele for um número inteiro.

    while True:
        try:
            print(f"\n{abs(int(n))} páginas, certo? ")
            return abs(int(n))
        except (ValueError, NameError, TypeError, KeyboardInterrupt):
            print(f'{Fore.RED}Entrada inválida. {Fore.RESET}', end='', flush=True)
            n = str(input('Digite o número de páginas que deseja ver: ').strip())


class BotAmazon:
    # Bot que acessa o site Amazon, procura produtos e preços e cria uma planilha automaticamente.

    def __init__(self):
        self.preco_com_decimal = list()
        self.lista_nome = list()
        self.pagina_atual = 1

        self.saudacao_e_pesquisa()
        self.iniciar_driver()
        self.acessar_site(self.pesquisa)
        self.varredura_do_site()
        self.criando_planilha()

    def saudacao_e_pesquisa(self):
        print('=' * 35)
        print(f'{Fore.MAGENTA}{Style.BRIGHT}')
        print('WEB SCRAPPING DO SITE "AMAZON"'.center(35))
        print(f'{Fore.RESET}{Style.RESET_ALL}')
        print('=' * 35)
        while True:
            self.pesquisa = str(input('O que deseja pesquisar no site Amazon.com.br? '))
            if self.pesquisa.strip() != '':
                break
        self.paginas = readint(str(input('Até quantas páginas deseja procurar? ').strip()))
        if self.paginas == 0:
            quit()
        print('Ao fim do processo, a página será fechada automaticamente.')
        sleep(2)
        print('Aguarde...')
        line()
        return self.pesquisa, self.paginas

    def iniciar_driver(self):
        try:
            service = Service('C:/PATH/TO/CHROMEDRIVER/chromedriver')  # Especifique AQUI o caminho para o chromedriver, com esse mesmo formato.
            self.driver = webdriver.Chrome(service=service)
        except Exception as ex:
            print(f"{Fore.RED}Erro ao tentar iniciar o Chromedriver.{Fore.RESET}\n"
                  f"Verifique se especificou o caminho para o chromedriver corretamente no método 'iniciar_driver'.\n"
                  f"Em caso positivo, veja se a versão é"
                  f" compatível com sua versão do Google Chrome.\nErro do tipo: {type(ex).__name__}")
            quit()

    def acessar_site(self, pesquisa):
        self.driver.get('https://www.amazon.com.br')
        search_box = self.driver.find_element(By.ID, 'twotabsearchtextbox')
        search_box.click()
        search_box.send_keys(pesquisa)
        self.driver.find_element(By.ID, 'nav-search-submit-button').click()

    def varredura_do_site(self):
        while True:
            try:
                # Programa aguarda até o primeiro elemento-nome carregar. Caso não carregue, o programa encerra:

                print(f'{Fore.GREEN}Extraindo dados da página {self.pagina_atual}.{Fore.RESET}')

                WebDriverWait(self.driver, 20).until(
                    ec.presence_of_element_located(
                        (By.XPATH, '//div[@class="s-main-slot s-result-list s-search-results sg-row"]'
                                   '//*[contains(@data-component-type, "s-search-result")]')))
                sleep(3)
            except Exception as ex:
                self.driver.quit()
                print(f'{Fore.RED}Erro ao carregar a página{Fore.RESET}.\nTipo de erro: {type(ex).__name__}')
                sys.exit(1)

            # Comando que busca o número identificador da posição do primeiro elemento.
            # Depois de guardar o número na variável "index", todos os nomes de produtos
            # do site serão localizados por iteração em um loop while:

            index = self.driver.find_element(
                By.XPATH, '//div[@class="s-main-slot s-result-list s-search-results sg-row"]//*[cont'
                          'ains(@data-component-type, "s-search-result")]').get_attribute('data-index')
            index = int(index)

            while True:
                try:
                    # Procura o nome do produto na página. Se não encontre nenhum, "break" e o programa segue.

                    nome = self.driver.find_element(By.XPATH, f'//div[@data-index= "{index}"]//*[contains(@class, "a-si'
                                                    f'ze-mini a-spacing-none a-color-base")]')
                except (NoSuchAttributeException, NoSuchElementException, KeyboardInterrupt):
                    break

                try:
                    # Verifica se o produto possui preço. Se houver, o produto será adicionado à planilha.
                    # Caso contrário, será ignorado e não irá para a planilha:

                    preco_inteiro = self.driver.find_element(
                        By.XPATH, f'//div[@data-index= "{index}"]//span[@class="a-price-whole"]')
                    preco_fracao = self.driver.find_element(
                        By.XPATH, f'//div[@data-index= "{index}"]//span[@class="a-price-fraction"]')
                except (NoSuchAttributeException, NoSuchElementException, KeyboardInterrupt):
                    pass
                else:
                    self.lista_nome.append(nome.text)
                    self.preco_com_decimal.append(f'{preco_inteiro.text},{preco_fracao.text}')
                finally:
                    index += 1

            for a in range(0, 2):
                # Adiciona dois espaços vazios que serão linhas vazias na planilha. Fica mais organizado :D

                self.lista_nome.append('')
                self.preco_com_decimal.append('')

            try:
                # Tenta localizar o botão "Próximo". Caso não tenha, o navegador fecha e a criação de planilha  começará

                if self.pagina_atual == self.paginas:
                    self.driver.quit()
                    break

                WebDriverWait(self.driver, 15).until(ec.element_to_be_clickable((
                    By.XPATH, '//a[@class="s-pagination-item s-pagination-next s-pagination-button s-pagination-separator"]')))
                sleep(2)
                botao_proximo = self.driver.find_element(By.XPATH, '//a[@class="s-pagination-item s-pagination-next s-'
                                                                   'pagination-button s-pagination-separator"]')
                botao_proximo.click()
            except (NoSuchAttributeException, NoSuchElementException, TimeoutException, KeyboardInterrupt):
                self.driver.quit()
                print('Não há mais páginas para acessar.')
                break
            else:
                self.pagina_atual += 1

        line()
        print(f'Foram extraídas informações de {self.pagina_atual} páginas.')
        line()

    def criando_planilha(self):
        print(f'{Fore.CYAN}Criando planilha', end='', flush=True)
        for a in range(0, 3):
            print('.', end='', flush=True), sleep(1)
        print(f'{Fore.RESET}')

        hoje_data = f'{datetime.today().day}.{datetime.today().month}.{datetime.today().year}'
        hoje_horario = f'{datetime.today().hour}h{datetime.today().minute}min'
        name = f"Planilha preços Amazon {hoje_data} ({hoje_horario}).xlsx"

        wb = xlsxwriter.Workbook(name)
        planilha = wb.add_worksheet(self.pesquisa.title())
        planilha.write('A1', 'PRODUTO')
        planilha.write('B1', 'PREÇO (R$)')

        index = 3

        for a in range(0, len(self.lista_nome)):
            planilha.write(index, 0, self.lista_nome[index - 3])
            planilha.write(index, 1, self.preco_com_decimal[index - 3])
            index += 1
        wb.close()

        print(f'{Fore.MAGENTA}"{name}" criada com sucesso!{Fore.RESET}')
        line()
        print('Até logo :D'.center(42))


bot = BotAmazon()

# Bot criado por: Rwann Pabblo.
