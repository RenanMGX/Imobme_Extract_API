"""
WARNING:

Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the dependencies.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at
https://documentation.botcity.dev/tutorials/custom-automations/python-custom/
"""

# Import for integration with BotCity Maestro SDK
from botcity.maestro import * #type: ignore
import traceback
from patrimar_dependencies.gemini_ia import ErrorIA
from patrimar_dependencies.screenshot import screenshot
from Entities.processos import Processos
from main import ExecuteAPP
import json
import os

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False #type: ignore


class Execute:
    @staticmethod
    def start():            
        crd_param = execution.parameters.get("crd")
        if not isinstance(crd_param, str):
            raise ValueError("Parâmetro 'crd_param' deve ser uma string representando o label da credencial.")
        
        lista_relatorios_param = execution.parameters.get("lista_relatorios")
        if not lista_relatorios_param:
            raise Exception(f"O parametro {lista_relatorios_param=} está vazio!")
        else:
            temp_lista_relatorios = str(lista_relatorios_param)
            lista_relatorios:dict = {value.split(',')[0]:{"file_name": value.split(',')[1]} for value in temp_lista_relatorios.split(';')}

        destino_param = execution.parameters.get("destino")
        if not destino_param:
            raise Exception(f"O parametro {destino_param=} está vazio!")
        else:
            destino:str = str(destino_param)
            for key, value in lista_relatorios.items():
                lista_relatorios[key]["destino"] = destino
        
        extension_param = execution.parameters.get("extension")
        if not extension_param:
            raise Exception(f"O parametro {extension_param=} está vazio!")
        else:
            extension:str = str(extension_param)
            for key, value in lista_relatorios.items():
                lista_relatorios[key]["extension"] = extension
                
        headless_param = execution.parameters.get("headless")
        if not headless_param:
            headless = False
        else:
            headless = str(headless_param).lower() == "true"
            
        quantidade_param = execution.parameters.get("quantidade")
        if not quantidade_param:
            quantidade = 1
        else:
            try:
                quantidade = int(str(quantidade_param))
            except:
                quantidade = 1
        
        p.total = len(lista_relatorios)
        
        app = ExecuteAPP(
            login=maestro.get_credential(label=crd_param, key="login"),
            password=maestro.get_credential(label=crd_param, key="password"),
            url=maestro.get_credential(label=crd_param, key="url"),
            headless=headless,
            p=p
        )
            
        try:
            app.start(lista_relatorios=lista_relatorios, quantidade=quantidade)
            
            if os.path.exists(app.path_api):
                api_path = os.listdir(app.path_api)
                if api_path:
                    for file in api_path:
                        file = os.path.join(app.path_api, file)
                        maestro.post_artifact(
                            task_id=int(execution.task_id),
                            artifact_name=os.path.basename(file),
                            filepath=file
                        ) 
                    
        finally:    
            app._limpar()

        p.add_processado()

if __name__ == '__main__':
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()
    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    task_name = execution.parameters.get('task_name')
    
    p = Processos(1)

    try:
        Execute.start()
        
        maestro.finish_task(
                    task_id=execution.task_id,
                    status=AutomationTaskFinishStatus.SUCCESS,
                    message=f"Tarefa {task_name} finalizada com sucesso",
                    total_items=p.total, # Número total de itens processados
                    processed_items=p.processados, # Número de itens processados com sucesso
                    failed_items=p.falhas # Número de itens processados com falha
        )
        
    except Exception as error:
        ia_response = "Sem Resposta da IA"
        try:
            token = maestro.get_credential(label="GeminiIA-Token-Default", key="token")
            if isinstance(token, str):
                ia_result = ErrorIA.error_message(
                    token=token,
                    message=traceback.format_exc()
                )
                ia_response = ia_result.replace("\n", " ")
        except Exception as e:
            maestro.error(task_id=int(execution.task_id), exception=e)

        maestro.error(task_id=int(execution.task_id), exception=error, screenshot=screenshot(), tags={"IA Response": ia_response})
        maestro.finish_task(
                    task_id=execution.task_id,
                    status=AutomationTaskFinishStatus.FAILED,
                    message=f"Tarefa {task_name} finalizada com Error",
                    total_items=p.total, # Número total de itens processados
                    processed_items=p.processados, # Número de itens processados com sucesso
                    failed_items=p.falhas # Número de itens processados com falha
        )
