import os
import xlwings as xw
import pandas as pd
import pythoncom
import json
import shutil

from patrimar_dependencies.functions import Functions
from typing import Literal

class Arquivos:
    @property
    def origin_file_path(self) -> str:
        return self.__origin_file_path
    
    @property
    def df(self) -> pd.DataFrame:
        return self.__df
    
    def __init__(self, file_path:str) -> None:
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"O arquivo {file_path} não foi encontrado.")
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"O arquivo {file_path} não foi encontrado.")
        
        if not file_path.lower().endswith(('.xls', '.xlsx', '.xlsm')):
            raise ValueError("O arquivo deve ser uma planilha do Excel com extensão .xls, .xlsx ou .xlsm")
        
        pythoncom.CoInitialize()
        try:
            app = xw.App(visible=False)
            with app.books.open(file_path) as wb:
                if 'Parâmetros' in wb.sheet_names:
                    wb.sheets['Parâmetros'].delete()
                wb.save()
            app.quit()
        finally:
            pythoncom.CoUninitialize()
        
        Functions().fechar_excel(file_path)
        
        self.__origin_file_path = file_path
        self.__df = pd.read_excel(file_path)
        
    def save_json_to(
        self, 
        move_to:str,
        *,
        file_name:str="",
        orient:Literal['columns', 'index', 'records', 'split', 'table', 'values'] = 'records',
        date_format:Literal['iso', 'epoch'] = 'iso',        
    ):
        if not os.path.isdir(move_to):
            raise NotADirectoryError(f"O diretório {move_to} não foi encontrado.")
        if not os.path.exists(move_to):
            raise NotADirectoryError(f"O diretório {move_to} não foi encontrado.")
        
        if file_name:
            if file_name == "DEFAULT":
                temp_name = os.path.basename(self.origin_file_path)
                file_name = temp_name.split('_')[0] + ".json"
            if not file_name.lower().endswith('.json'):
                file_name += '.json'
            move_to = os.path.join(move_to, file_name)
        else:
            temp_name = os.path.basename(self.origin_file_path)
            temp_name = os.path.splitext(temp_name)[0] + ".json"
            move_to = os.path.join(move_to, temp_name)
        
        self.df.to_json(move_to, orient=orient, date_format=date_format)
        
    def save_csv_to(
        self,
        move_to:str,
        *,
        file_name:str="",
        sep:str=";",
        index:bool=False,
        encoding:Literal['utf-8', 'latin1', 'utf-16', 'utf-32', 'ascii'] = 'latin1',
        errors:Literal['backslashreplace', 'ignore', 'namereplace', 'replace', 'strict', 'surrogateescape', 'xmlcharrefreplace'] = 'ignore',
        decimal:Literal['.', ','] = ','
    ):
        if not os.path.isdir(move_to):
            raise NotADirectoryError(f"O diretório {move_to} não foi encontrado.")
        if not os.path.exists(move_to):
            raise NotADirectoryError(f"O diretório {move_to} não foi encontrado.")
        
        if file_name:
            if file_name == "DEFAULT":
                temp_name = os.path.basename(self.origin_file_path)
                file_name = temp_name.split('_')[0] + ".csv"
            if not file_name.lower().endswith('.csv'):
                file_name += '.csv'
            move_to = os.path.join(move_to, file_name)
        else:
            temp_name = os.path.basename(self.origin_file_path)
            temp_name = os.path.splitext(temp_name)[0] + ".csv"
            move_to = os.path.join(move_to, temp_name)
        
        self.df.to_csv(move_to, sep=sep, index=index, encoding=encoding, errors=errors, decimal=decimal)

    def save_excel_to(
        self,
        move_to:str,
        *,
        file_name:str="",
    ):
        if not os.path.isdir(move_to):
            raise NotADirectoryError(f"O diretório {move_to} não foi encontrado.")
        if not os.path.exists(move_to):
            raise NotADirectoryError(f"O diretório {move_to} não foi encontrado.")
        
        if file_name:
            if file_name == "DEFAULT":
                temp_name = os.path.basename(self.origin_file_path)
                file_name = temp_name.split('_')[0] + ".xlsx"
            if not file_name.lower().endswith('.xlsx'):
                file_name += '.xlsx'
            move_to = os.path.join(move_to, file_name)
        else:
            temp_name = os.path.basename(self.origin_file_path)
            temp_name = os.path.splitext(temp_name)[0] + ".xlsx"
            move_to = os.path.join(move_to, temp_name)
        
        shutil.copy2(self.origin_file_path, move_to)
        
if __name__ == "__main__":
    pass
      