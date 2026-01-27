import os
import shutil
import re
import string
import random

from patrimar_dependencies.navegador_chrome import NavegadorChrome, By, Keys, WebDriver, WebElement
from patrimar_dependencies.functions import P
from typing import Literal, List
from functools import wraps
from time import sleep
from .exeptions import *
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from botcity.maestro import * #type: ignore

maestro:BotMaestroSDK|None = BotMaestroSDK.from_sys_args()
try:
    execution = maestro.get_execution()
except:
    maestro = None


class Imobme(NavegadorChrome):
    @property
    def login(self) -> str:
        return self.__login
    
    @property
    def password(self) -> str:
        return self.__password
    
    @property
    def url(self) -> str:
        return self.__url
    
    @property
    def base_url(self) -> str:
        if (url:=re.search(r'[A-z]+://[A-z0-9.]+/', (self.url if self.url.endswith('/') else self.url+'/'))) is not None:
            return url.group()
        raise UrlError("URL inválida!")
    
    def _find_element_especial(self, by=By.ID, value = None, *, timeout = 10, force = False, wait_before = 0, wait_after = 0):
        """
            É para ser utilizado apenas no decorador de login
        """
        return super().find_element(by, value, timeout=timeout, force=force, wait_before=wait_before, wait_after=wait_after)
    
    @staticmethod
    def logar(f):
        @wraps(f)
        def wrap(*args, **kwargs):
            self:Imobme = args[0]
            
            tela_de_login = False
            
            try:
                sleep(0.5)
                self.find_element_native(By.ID, "login")
                tela_de_login = True
            except:
                pass
            
            if tela_de_login:
                self._find_element_especial(By.ID, "login").send_keys(self.login)
                self._find_element_especial(By.ID, "password").send_keys(self.password)
                self._find_element_especial(By.ID, "password").send_keys(Keys.RETURN)
                
                try:
                    if self._find_element_especial(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/form/div/ul/li', force=True, timeout=2).text == 'Login não encontrado.':
                        raise PermissionError("Login não encontrado.")
                except PermissionError as error:
                    raise error
                except Exception:
                    pass
                
                
                try:
                    if 'Senha Inválida.' in (return_error:=self._find_element_especial(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/form/div/ul/li', force=True, timeout=2).text):
                        raise PermissionError(return_error)
                except PermissionError as error:
                    raise error
                except Exception:
                    pass
                
                
                try:
                    self._find_element_especial(By.XPATH, '/html/body/div[2]/div[3]/div/button[1]/span', force=True, timeout=2).click()
                except:
                    pass

                print("Login realizado com sucesso.")
           
            return f(*args, **kwargs)
        return wrap

    @logar
    def _load_page(self, endpoint:str):
        if not endpoint.endswith('/'):
            endpoint += '/'
        if endpoint.startswith('/'):
            endpoint = endpoint[1:]
        
        url = os.path.join(self.base_url, endpoint)
        print(P(f"Carregando página: {url}...          ", color='yellow'))  
        self.get(url)
    
    @logar   
    def _select_relatorio(self, relat:str):
        self._load_page("Relatorio")
        self.maximize_window()
        self.find_element(By.XPATH, '//*[@id="Content"]').location_once_scrolled_into_view
        self.find_element(By.ID, 'Relatorios_chosen').click() # clique em selecionar Relatorios
        
        ul_tag = self.find_element(By.XPATH, '//*[@id="Relatorios_chosen"]/div/ul')
        
        for tag_li in ul_tag.find_elements(By.TAG_NAME, 'li'):
            try:
                if tag_li.text.lower().replace(" ", "") == relat.lower().replace(" ", ""):
                    tag_li.click()
                    return
            except Exception as error:
                pass
        
        raise RelatorioNotFound(f"relatorio {relat} não foi encontrado na lista de relatorios")

    def _verificar_download(self):
        sleep(5)
        for _ in range(30*60):
            arquivos = os.listdir(self.download_path)
            arquivos_caminhos = [os.path.join(self.download_path, f) for f in arquivos]
            file = max(arquivos_caminhos, key=os.path.getmtime)
            if not file.endswith('.crdownload'):
                sleep(3)
                return file
            sleep(1)
        raise TimeoutError("O download demorou mais de 30 minutos para ser concluído.")
            
    def __del__(self):
        self._encerrar()
            
    def _encerrar(self):
        # if not self.debug:
        #     try:
        #         shutil.rmtree(self.download_path)
        #     except Exception as error:
        #         pass
        #         #print(P(f"não foi possivel apagar a pasta {self.download_path} pelo motivo:\n{str(error)}", color='red'))               
        try:
            self.quit()
        except:
            pass  
    
    def __init__(
        self, *,
        login:str,
        password:str,
        url:str,
        speak:bool = False, 
        download_path:str = os.path.join(os.getcwd(), "downloads", datetime.now().strftime(f"relatorios_%Y%m%d_%H%M%S_{''.join(random.choices(string.ascii_letters + string.digits, k=16))}")), 
        headless:bool = False,
        clear_download_folder:bool = True,
        debug:bool = False
    ):
        self.debug = debug
        if os.path.isfile(download_path):
            download_path = os.path.dirname(download_path)
            
        self.download_path = download_path
        self.__login:str = login
        self.__password:str = password        
        self.__url:str = url
        
        if not os.path.exists(self.download_path):
            os.makedirs(self.download_path)
        else:
            if clear_download_folder:
                shutil.rmtree(self.download_path)
                os.makedirs(self.download_path)
        
        super().__init__(speak=speak, download_path=self.download_path, headless=headless)
        self.get(self.base_url)
        self.get(self.base_url)
        
    def clear_download_folder(self):
        if os.path.exists(self.download_path):
            shutil.rmtree(self.download_path)
            os.makedirs(self.download_path)
            
    @logar
    def find_element(self, by=By.ID, value = None, *, timeout = 10, force = False, wait_before = 0, wait_after = 0):
        return super().find_element(by, value, timeout=timeout, force=force, wait_before=wait_before, wait_after=wait_after)
    
    @logar
    def limpar_relatorios(self):
        self._load_page("Relatorio")
        cont_errors = 0
        linha = 1
        for _ in range(1000):   
            try:
                self.title
            except:
                return
            if linha > 20:
                return        
            try:
                self.find_element(By.XPATH, f'//*[@id="result-table"]/tbody/tr[{linha}]/td[11]/a/i', timeout=5)
            except:
                linha += 1
                            
            try:
                
                self.find_element(By.XPATH, f'//*[@id="result-table"]/tbody/tr[{linha}]/td[12]/a/i', timeout=2).click()
                
                self.find_element(By.XPATH, '/html/body/div[5]/div[3]/div/button[1]/span', timeout=2).click()
            except:
                try:
                    self.find_element(By.XPATH, '/html/body/div[5]/div[3]/div/button[2]/span', timeout=2).click()
                except:
                    pass
                try:
                    self.find_element(By.XPATH, f'//*[@id="result-table"]/tbody/tr[{linha}]/td[12]/a/i', timeout=2).click()
                except:
                    continue
                try:
                    self.find_element(By.XPATH, '/html/body/div[5]/div[3]/div/button[1]/span', timeout=2).click()
                except:
                    continue
                    
                #cont_errors += 1
                print(f"Erro: {cont_errors=}")
                if cont_errors > 3:
                    print("Erro: Excedeu o limite de tentativas")
                    break    
    
    valid_relatorios = [
            "imobme_empreendimento",
            "imobme_controle_vendas",
            "imobme_controle_vendas_90_dias",
            "imobme_contratos_rescindidos",
            "imobme_contratos_rescindidos_90_dias",
            "imobme_dados_contrato",
            "imobme_previsao_receita",
            "imobme_relacao_clientes",
            "imobme_relacao_clientes_x_clientes",
            "imobme_cadastro_datas",
            "recebimentos_compensados",
            "imobme_controle_estoque",
        ]
    
    @logar
    def extrair_relatorios(self, relatorios:List[str]) -> List[dict]:
        """Lista de relatórios a serem extraídos:
        - imobme_empreendimento
        - imobme_controle_vendas
        - imobme_controle_vendas_90_dias
        - imobme_contratos_rescindidos
        - imobme_contratos_rescindidos_90_dias
        - imobme_dados_contrato
        - imobme_previsao_receita
        - imobme_relacao_clientes
        - imobme_relacao_clientes_x_clientes
        - imobme_cadastro_datas
        - recebimentos_compensados
        - imobme_controle_estoque

        Args:
            relatorios (List[str]): lista dos relatórios a serem extraídos
        """
        self._load_page("Relatorio")
        
        # Extração de Relatorios
        ids_relatorios:dict = {}
        
        ### --------------------------------------------------------- ###
        # Relatorio IMOBME - Empreendimento
        if "imobme_empreendimento" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Empreendimento")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em selecionar Emprendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clique em selecionar todos os empreendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em selecionar Emprendimentos novamente para sair
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_empreendimento"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_empreendimento' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)
        
        # Relatorio IMOBME - Controle de Vendas          
        if "imobme_controle_vendas" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Controle de Vendas")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="DataInicio"]').send_keys("01012015") # escreve a data de inicio padrao 01/01/2015
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data hoje
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="TipoDataSelecionada_chosen"]/a').click() # clique em Tipo Data
                    self.find_element(By.XPATH, '//*[@id="TipoDataSelecionada_chosen_o_0"]').click() # clique em Data Lançamento Venda
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em Empreendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clique em Todos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em Empreendimentos
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_controle_vendas"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_controle_vendas' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)
                    
        # Relatorio IMOBME - Controle de Vendas 90 dias     
        if "imobme_controle_vendas_90_dias" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Controle de Vendas")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="DataInicio"]').send_keys((datetime.now() - relativedelta(days=90)).strftime("%d%m%Y")) # escreve a data de inicio com um range de 90 dias
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data hoje
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="TipoDataSelecionada_chosen"]/a').click() # clique em Tipo Data
                    self.find_element(By.XPATH, '//*[@id="TipoDataSelecionada_chosen_o_0"]').click() # clique em Data Lançamento Venda
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em Empreendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clique em Todos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em Empreendimentos
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_controle_vendas_90_dias"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_controle_vendas_90_dias' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)
                    
        # Relatorio IMOBME - Contratos Rescindidos 90 dias     
        if "imobme_contratos_rescindidos" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Contratos Rescindidos")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="DataInicio"]').send_keys("01012015") # escreve a data de inicio padrao 01/01/2015
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data hoje
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[2]/div/div[3]/div/button').click()  # clique em Tipo de Contrato
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[2]/div/div[3]/div/ul/li[2]/a/label/input').click() # clique em Todos
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[2]/div/div[3]/div/button').click() # clique em Tipo de Contrato
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clique em Todos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_contratos_rescindidos"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_contratos_rescindidos' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)

        # Relatorio IMOBME - Contratos Rescindidos 90 dias     
        if "imobme_contratos_rescindidos_90_dias" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Contratos Rescindidos")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="DataInicio"]').send_keys((datetime.now() - relativedelta(days=90)).strftime("%d%m%Y")) # escreve a data de inicio padrao 01/01/2015
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data hoje
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[2]/div/div[3]/div/button').click()  # clique em Tipo de Contrato
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[2]/div/div[3]/div/ul/li[2]/a/label/input').click() # clique em Todos
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[2]/div/div[3]/div/button').click() # clique em Tipo de Contrato
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clique em Todos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_contratos_rescindidos_90_dias"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_contratos_rescindidos_90_dias' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)

        # Relatorio IMOBME - Dados do Contrato     
        if "imobme_dados_contrato" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Dados do Contrato")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label').click() # clica em todos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[1]/div').click() # clica fora
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[3]/div[1]/div/button').click() # clica em tipos de contrato
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[3]/div[1]/div/ul/li[2]/a/label/input').click() # clica em Todos
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[1]/div').click() # clica fora
                    self.find_element(By.XPATH, '//*[@id="DataBase"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data hoje
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_dados_contrato"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_dados_contrato' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)

        # Relatorio IMOBME - Previsão de Receita     
        if "imobme_previsao_receita" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Previsão de Receita")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="DataInicio"]').send_keys("01012015") # escreve a data de inicio padrao 01/01/2015
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="DataFim"]').send_keys((datetime.now() + relativedelta(years=25)).strftime("%d%m%Y")) # escreve a data de fim padrao com a data atual mais 25 anos
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clica em todos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                    self.find_element(By.XPATH, '//*[@id="DataBase"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data de hoje
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    sleep(3)
                    
                    self.find_element(By.ID, 'parametrosReport').location_once_scrolled_into_view
                    
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[4]/div/div[2]/div/button').click() # clica em Tipo Parcela
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[4]/div/div[2]/div/ul/li[2]/a/label/input').click() # clica em Todos
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[4]/div/div[2]/div/button').click() # clica em Tipo Parcela
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_previsao_receita"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_previsao_receita' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)
     
        # Relatorio IMOBME - Relação de Clientes    
        if "imobme_relacao_clientes" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Relação de Clientes")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="DataInicio"]').send_keys("01012015") # escreve a data de inicio padrao 01/01/2015
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data de hoje
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_relacao_clientes"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_relacao_clientes' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)
     
        # Relatorio IMOBME - Relação de Clientes x Clientes  
        if "imobme_relacao_clientes_x_clientes" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Relação de Clientes")
                    sleep(2)
                    
                    self.find_element(By.XPATH, '//*[@id="tipoReportCliente_chosen"]').click() # clica em tipo de relatorio
                    self.find_element(By.XPATH, '//*[@id="tipoReportCliente_chosen_o_1"]').click() # clica em clientes x contratos
                    
                    try:
                        for h4 in self.find_elements(By.TAG_NAME, 'h4'):
                            if h4.text == "Relatórios":
                                h4.click()
                                h4.location_once_scrolled_into_view
                                break
                    except:
                        pass
                    
                    self.find_element(By.XPATH, '//*[@id="dvTipoContrato"]/div/button').click() # clica em tipo de contrato
                    self.find_element(By.XPATH, '//*[@id="dvTipoContrato"]/div/ul/li[2]/a/label/input').click() # clica em todos
                    self.find_element(By.XPATH, '//*[@id="dvTipoContrato"]/div/button').click() # clica em tipo de contrato
                    
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clica em todos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                    self.find_element(By.XPATH, '//*[@id="DataInicio"]').send_keys("01012015") # escreve a data de inicio padrao 01/01/2015
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data de hoje
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    
                    #import pdb; pdb.set_trace()
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_relacao_clientes_x_clientes"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_relacao_clientes_x_clientes' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)
                    
        # Relatorio IMOBME - Cadastro de Datas
        if "imobme_cadastro_datas" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Cadastro de Datas")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label').click() # clica em todos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_cadastro_datas"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_cadastro_datas' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)
         
        # Relatorio Recebimentos Compensados
        if "recebimentos_compensados" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("Recebimentos Compensados")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="DataInicio"]').send_keys("01012020") # escreve a data de inicio padrao 01/01/2020
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data de hoje
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label').click() # clica em todos
                    self.find_element(By.XPATH, '//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["recebimentos_compensados"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'recebimentos_compensados' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)
         
        # Relatorio IMOBME - Controle de Estoque
        if "imobme_controle_estoque" in relatorios:
            for _ in range(5):
                try:
                    self._select_relatorio("IMOBME - Controle de Estoque")
                    sleep(.5)
                    
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[2]/div[1]/div/div/button').click() # clica em Empreendimentos
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[2]/div[1]/div/div/ul/li[2]/a/label/input').click() # clica em todos
                    self.find_element(By.XPATH, '//*[@id="parametrosReport"]/div[2]/div[1]/div/div/button').click() # clica em Empreendimentos novamente para sair
                    self.find_element(By.XPATH, '//*[@id="DataBase"]').send_keys(datetime.now().strftime("%d%m%Y"))
                    self.find_element(By.XPATH, '//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                    
                    self.find_element(By.XPATH, '//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                    sleep(7)
                    ids_relatorios["imobme_controle_estoque"] = self.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                    print(P(f"o relatorio 'imobme_controle_estoque' foi gerado!"))
                    break
                    
                except Exception as error:
                    print(P(str(error), color='red'))
                    if _ >= 4:
                        if maestro:
                            maestro.error(task_id=int(execution.task_id), exception=error)
                    
        ### --------------------------------------------------------- ###         
        
        #finalmente, se nenhum relatorio foi gerado, retorna uma lista vazia           
        if not ids_relatorios:
            print(P(f"Nenhum relatorio foi gerado!", color='red'))
            return []
                
        print(P(f"Aguardando o download de {len(ids_relatorios)} relatórios...", color='red'))
        
        # Aguardar e fazer download dos relatórios
        relatorios_paths = []
        
        cont_final: int = 0
        data_inicial = datetime.now()

        while True:
            try:
                self.title
            except:
                raise BrowserClosed("O navegador foi fechado inesperadamente.")
            
            if datetime.now() >= (timedelta(hours=1, minutes=30) + data_inicial):
                print()
                print(P("saida emergencia acionada a espera da geração dos relatorios superou as 1,5 horas", color='red'))
                #Logs().register(status='Report', description=f"saida emergencia acionada a espera da geração dos relatorios superou as 1,5 horas")
                raise TimeoutError("saida emergencia acionada a espera da geração dos relatorios superou as 1,5 horas")
            else:
                cont_final += 1
            if not ids_relatorios:
                break
            
            try:
                table = self.find_element(By.ID, 'result-table')
                tbody = table.find_element(By.TAG_NAME, 'tbody')
                for tr in tbody.find_elements(By.TAG_NAME, 'tr'):
                    for key,id in ids_relatorios.items():
                        downloaded = False
                        if id == tr.find_elements(By.TAG_NAME, 'td')[0].text:
                            for tag_a in tr.find_elements(By.TAG_NAME, 'a'):
                                if tag_a.get_attribute('title') == 'Download':
                                    print()
                                    print(P(f"o {id=} foi baixado!", color='green'))
                                    tag_a.send_keys(Keys.ENTER)
                                    #ids_relatorios.pop(ids_relatorios.index(id))
                                    del ids_relatorios[key]
                                    
                                    relatorios_paths.append({key:self._verificar_download()})
                                    downloaded = True
                                try:
                                    if downloaded:
                                        if tag_a.get_attribute('title') == 'Excluir':
                                            tag_a.send_keys(Keys.ENTER)
                                            #self.find_element(By.XPATH, '/html/body/div[5]/div[3]/div/button[1]/span').click()
                                            for button in self.find_elements(By.TAG_NAME, 'button'):
                                                if button.text == "Confirmar": #bt para confirmar exclusao
                                                    button.click()
                                                    break
                                            print()
                                            print(P(f"o {id=} foi excluido!", color='red'))
                                except:
                                    pass
                                try:
                                    #self.find_element(By.XPATH, '/html/body/div[5]/div[3]/div/button[2]/span').click()
                                    for button in self.find_elements(By.TAG_NAME, 'button'):
                                        if button.text == "Cancelar": #bt para cancelar exclusao
                                            button.click()
                                            break
                                except:
                                    pass
            except:
                sleep(5)
                continue
            
            self.find_element(By.ID, 'btnProximaDefinicao').click()
            sleep(5)
            print(P("Atualizando Pagina", color='yellow'), end='\r')
        print()
        
        return relatorios_paths
    
    @logar
    def teste(self):
        print("Iniciando teste...")
        self._load_page("Relatorio")
        import pdb; pdb.set_trace()
        self.find_element(By.TAG_NAME, "html")

if __name__ == "__main__":
    pass
