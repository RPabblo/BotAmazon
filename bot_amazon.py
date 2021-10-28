from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
from time import sleep
from colorama import Fore, Style
from datetime import datetime


def line():
    # Função que mostra na tela uma linha. Apenas um detalhe :D
    print('–' * 32)


def readint(n):
    # Função que recebe um input do usuário e retorna um valor somente se ele for um número inteiro.

    while True:
        try:
            return int(abs(n))
        except (ValueError, NameError, TypeError, KeyboardInterrupt):
            print(f'{Fore.RED}Entrada inválida. ', end='', flush=True)
            n = str(input('Digite o número de páginas que deseja ver: ').strip())


class BotAmazon:
    # Bot que acessa o site Amazon, procura produtos e preços e cria uma planilha automaticamente.

    def __init__(self):
        self.preco_com_decimal = list()
        self.lista_nome = list()
        self.pagina_atual = 1

        self.saudacao_e_pesquisa()
        self.driver = webdriver.Chrome('C://Users/Carlos/chromedriver')
        self.driver.maximize_window()
        self.acessar_site(self.pesquisa)
        self.varredura_do_site()
        self.criando_planilha()

    def saudacao_e_pesquisa(self):
        print('=' * 32)
        print(f'{Fore.MAGENTA}{Style.BRIGHT}')
        print('WEB SCRAPPING DO SITE "AMAZON"'.center(32))
        print(f'{Fore.RESET}{Style.RESET_ALL}')
        print('=' * 32)
        while True:
            self.pesquisa = str(input('O que deseja pesquisar no site Amazon.com.br? '))
            if self.pesquisa.strip() != '':
                break
        self.paginas = readint(str(input('Até quantas páginas deseja procurar? ').strip()))
        print('Ao fim do processo, a página será fechada automaticamente.')
        sleep(1)
        print('Aguarde...')
        line()
        return self.pesquisa, self.paginas

    def acessar_site(self, pesquisa):
        try:
            self.driver.get('https://www.amazon.com.br')
        except:
            print(f'{Fore.RED}Erro ao carregar o site.')
            self.driver.quit()
            quit()
        search_box = self.driver.find_element_by_id('twotabsearchtextbox')
        search_box.click()
        search_box.send_keys(pesquisa)
        self.driver.find_element_by_id('nav-search-submit-button').click()

    def varredura_do_site(self):
        while True:
            print(f'{Fore.GREEN}Extraindo dados da página {self.pagina_atual}.{Fore.RESET}')
            try:
                WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.XPATH,
                                                                                 '//div[@class="s-main-slot s-result-li'
                                                                                 'st s-search-results sg-row"]/div[2]//'
                                                                                 '*[contains(@class, "a-size-mini")]')))
                sleep(3)
            except:
                self.driver.quit()
                print(f'{Fore.RED}Erro ao carregar a página.{Fore.RESET}')
                break
            else:

                index = 1
                while True:
                    try:
                        # Procura o nome do produto na página. Se não encontrar nenhum, acontece um "break"
                        # e o programa segue.

                        nome = self.driver.find_element_by_xpath(
                                f'//div[@class="s-main-slot s-result-list s-search-results sg-row"]/div'
                                f'[{index}]//*[contains(@class, "a-size-mini")]')
                    except:
                        if index == 1:
                            index += 1
                        else:
                            break

                    else:
                        # Verifica se o produto possui preço
                        # Caso não, será ignorado e não irá para a planilha
                        try:
                            preco_inteiro = self.driver.find_element_by_xpath(
                                    f'//div[@class="s-main-slot s-result-list s-search-results sg-row"]/div['
                                    f'{index}]//span[@class="a-price-whole"]')
                            preco_fracao = self.driver.find_element_by_xpath(
                                    f'//div[@class="s-main-slot s-result-list s-search-results sg-row"]/div'
                                    f'[{index}]//span[@class="a-price-fraction"]')
                        except:
                            pass
                        else:
                            self.lista_nome.append(nome.text)
                            self.preco_com_decimal.append(f'{preco_inteiro.text},{preco_fracao.text}')
                        finally:
                            index += 1

                for a in range(0, 2):
                    self.lista_nome.append('')
                    self.preco_com_decimal.append('')

                # Tenta localizar o botão "Próximo".
                # Caso não tenha, o navegador fecha a criação de planilha  começará:
                try:
                    if self.pagina_atual == self.paginas:
                        self.driver.quit()
                        break
                    else:
                        WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.XPATH,
                                                                                         '//li[@class="a-last"]/a')))
                        sleep(2)
                        self.driver.find_element_by_xpath('//li[@class="a-last"]/a').click()
                except:
                    self.driver.quit()
                    print('Não há mais páginas para acessar.')
                    line()
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
        print('Até logo :D'.center(32))


bot = BotAmazon()

# Bot criado por: Rwann Pabblo.
