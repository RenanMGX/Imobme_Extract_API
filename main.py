import multiprocessing as mp
import sys
import signal
import os
import shutil
import random
import string

from Entities.imobme import Imobme
from Entities.arquivos import Arquivos
from patrimar_dependencies.functions import P
from typing import List, Dict
from patrimar_dependencies.sharepointfolder import SharePointFolders
from datetime import datetime
from botcity.maestro import * #type: ignore
from Entities.processos import Processos

maestro:BotMaestroSDK|None = BotMaestroSDK.from_sys_args()
try:
    execution = maestro.get_execution()
except:
    maestro = None


lock = mp.Lock()
def init_worker(l):
    global lock
    lock = l

class ExecuteAPP:
    def _separar_relatorios(self, *, lista:list, quantidade:int=4):
        li_final = []
        while lista:
            temp = []
            if len(lista) > quantidade:
                for _ in range(quantidade):
                    temp.append(lista.pop(0))
                li_final.append(temp)
            else:
                for _ in range(len(lista)):
                    temp.append(lista.pop(0))
                li_final.append(temp)
                    
        return li_final        
        
    def _limpar(self):
        all_dir_paths = list(set(self.__all_dir_paths))
        if all_dir_paths:
            for dir in all_dir_paths:
                shutil.rmtree(dir, ignore_errors=True)
        
        if os.path.exists(self.path_api):
            shutil.rmtree(self.path_api, ignore_errors=True)

    def __init__(self, *, login:str, password:str, url:str, headless:bool=False, p:Processos|None=None):
        self.__login = login
        self.__password = password
        self.__url = url
        
        self.__all_dir_paths = []
        
        self.path_api = os.path.join(os.getcwd(), 'downloads',  datetime.now().strftime(f"api_%Y%m%d_%H%M%S_{''.join(random.choices(string.ascii_letters + string.digits, k=16))}"))
        if not os.path.exists(self.path_api):
            os.makedirs(self.path_api)
            
        self.headless:bool = headless
        
        self.p:Processos|None = p

    def _extrair_relatorios(self, grupos:list) -> List[dict]:
        global lock
        
        with lock:
            imob = Imobme(
                login=self.__login,
                password=self.__password,
                url=self.__url,
                headless=self.headless,
            )

        try:
            resultado = imob.extrair_relatorios(grupos)
        except Exception as err:
            if maestro:
                maestro.error(task_id=int(execution.task_id), exception=err)
            
            print(P((type(err), err), color='red'))
            resultado = []
        finally:
            imob._encerrar()
            del imob
            
        #all_dir_paths = []
        
        lista_relatorios_temp = {}
        
        if resultado:
            for relatorio in resultado:
                relatorio:dict
                for key, value in relatorio.items():
                    lista_relatorios_temp[key] = self.lista_relatorios[key]
                    lista_relatorios_temp[key]["path_file"] = value

        #self.__all_dir_paths = list(set(self.__all_dir_paths))
        
        for key, data in lista_relatorios_temp.items():
            data:dict
            if not "API:" in data['destino']:
                moveTo = data['destino'] if not "SHAREPOINT:" in data['destino'] else SharePointFolders(data['destino'].replace("SHAREPOINT:", "")).value
            else:
                moveTo = self.path_api
                        
            if data.get("extension") == "JSON":
                Arquivos(file_path=data['path_file']).save_json_to(move_to=moveTo, file_name=(data["file_name"] if data.get("file_name") else ""))
            elif data.get("extension") == "CSV":
                Arquivos(file_path=data['path_file']).save_csv_to(move_to=moveTo, file_name=(data["file_name"] if data.get("file_name") else ""))
            else:
                Arquivos(file_path=data['path_file']).save_excel_to(move_to=moveTo, file_name=(data["file_name"] if data.get("file_name") else ""))
        
        if not self.p is None:
            self.p.add_processado(len(lista_relatorios_temp))
            

        return resultado
    
    def limpar_relatorios(self) -> bool:
        imob = Imobme(
            login=self.__login,
            password=self.__password,
            url=self.__url,
            headless=self.headless,
        )
        try:
            imob.limpar_relatorios()
            return True
        except:
            return False
    
    
    @staticmethod
    def validar_lista_relatorios(f):
        def wrap(*args, **kwargs):            
            lista_relatorios = kwargs.get('lista_relatorios', None)
            if not lista_relatorios:
                raise ValueError("A lista de relatórios não pode estar vazia.")
            
            tem_relatorio = False
            for relat in Imobme.valid_relatorios:
                if relat in lista_relatorios:
                    tem_relatorio = True
                    break
                    
            if not tem_relatorio:
                raise ValueError(f"A lista de relatórios não contém nenhum relatório válido. relatorios validos: {Imobme.valid_relatorios}")
                
                            
            for relatorio, data in lista_relatorios.items():
                if "destino" in data:
                    if "SHAREPOINT:" in data["destino"]:
                        try:
                            SharePointFolders(data["destino"].replace("SHAREPOINT:", "")).value
                        except FileNotFoundError as error:
                            raise FileNotFoundError(f"O caminho do sharepoint para o relatório '{relatorio}' não foi encontrado. {str(error)}")
                    elif "API:" in data["destino"]:
                        pass
                    else:
                        if not os.path.isdir(data["destino"]):
                            raise NotADirectoryError(f"O diretório de destino para o relatório '{relatorio}' não foi encontrado.\n{data['destino']}")
                        if not os.path.exists(data["destino"]):
                            raise NotADirectoryError(f"O diretório de destino para o relatório '{relatorio}' não foi encontrado.\n{data['destino']}")
                else:
                    raise ValueError(f"O destino de extração para o relatório '{relatorio}' não foi especificado.")
            

            #raise Exception("Tudo Valido")
            return f(*args, **kwargs)
        return wrap

    @validar_lista_relatorios
    def start(self, *, lista_relatorios:Dict[str, dict], quantidade:int=1):
        """Estrutura do dicionario
            {
                "nome_relatorio": {
                    "path_file": "caminho/do/relatorio.xlsx", <---- é preenchido pelo proprio script
                    "destino": "caminho/para/salvar/o/arquivo", <---- caminho bruto de onde sera salvo | "SHAREPOINT:caminho/para/salvar/o/arquivo/no/sharepoint" <---- caminho apenas do endereço final do sharepoint, "API:" <---- retornar via api
                    "extension": "XLSX" | "CSV" | "JSON"
                    "file_name": "nome_do_arquivo" <---- opçoes
                }
            }
        

        Args:
            lista_relatorios (Dict[str, dict]): _description_
        """
        
        self.lista_relatorios = lista_relatorios
        
        num_process = mp.cpu_count() - 1 if len(lista_relatorios) > (mp.cpu_count() - 1) else len(lista_relatorios)
        
        lista_relatorios_temp = self._separar_relatorios(lista=list(lista_relatorios.keys()), quantidade=quantidade) #type:ignore
        
        l = mp.Lock()
        pool = mp.Pool(initializer=init_worker, initargs=(l,),processes=num_process)
        
        resultados:List[dict] = pool.map(self._extrair_relatorios, lista_relatorios_temp) #type:ignore
        
        pool.close()
        pool.join()
        
        all_dir_paths = []
        if resultados:
            for i, resultado in enumerate(resultados):
                for relatorio in resultado:
                    relatorio:dict
                    for key, value in relatorio.items():
                        #lista_relatorios[key]["path_file"] = value
                        all_dir_paths.append(os.path.dirname(value))

        self.__all_dir_paths = list(set(all_dir_paths))
        
        
        # for key, data in lista_relatorios.items():
        #     data:dict
        #     if not "API:" in data['destino']:
        #         moveTo = data['destino'] if not "SHAREPOINT:" in data['destino'] else SharePointFolders(data['destino'].replace("SHAREPOINT:", "")).value
        #     else:
        #         moveTo = self.path_api
                        
        #     if data.get("extension") == "JSON":
        #         Arquivos(file_path=data['path_file']).save_json_to(move_to=moveTo, file_name=(data["file_name"] if data.get("file_name") else ""))
        #     elif data.get("extension") == "CSV":
        #         Arquivos(file_path=data['path_file']).save_csv_to(move_to=moveTo, file_name=(data["file_name"] if data.get("file_name") else ""))
        #     else:
        #         Arquivos(file_path=data['path_file']).save_excel_to(move_to=moveTo, file_name=(data["file_name"] if data.get("file_name") else ""))
                

    
if __name__ == "__main__":
    lista_relatorios = {
        "imobme_empreendimento": {
            "destino": r"C:\Users\renan.oliveira\Downloads\preparativo",
            "file_name":"DEFAULT",
            "extension": "JSON",
            #"path_file" : r"#material\Empreendimentos_48954_20250902-171050.xlsx"
        },
        "imobme_relacao_clientes_x_clientes": {
            "destino": r"C:\Users\renan.oliveira\Downloads\preparativo",
            "file_name":"DEFAULT",
            "extension": "JSON",
            #"path_file" : r"#material\Empreendimentos_48954_20250902-171050.xlsx"
        },
        "imobme_previsao_receita": {
            "destino": r"C:\Users\renan.oliveira\Downloads\preparativo",
            "file_name":"DEFAULT",
            "extension": "JSON",
            #"path_file" : r"#material\Empreendimentos_48954_20250902-171050.xlsx"

        }
        
    }
    
    
    from patrimar_dependencies.credenciais import Credential
    
    
    crd:dict = Credential(
        path_raiz=SharePointFolders(r'RPA - Dados\CRD\.patrimar_rpa\credenciais').value,
        name_file="IMOBME_PRD"
    ).load()
    
    app = ExecuteAPP(
        login=crd["login"],
        password=crd["password"],
        url=crd["url"],

    )
    try:
        app.start(lista_relatorios=lista_relatorios, quantidade=1)
    
        #input("parado aqui: ")
    finally:
        app._limpar()
    