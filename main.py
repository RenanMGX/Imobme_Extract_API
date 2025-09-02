import multiprocessing as mp
import sys
import signal

from Entities.imobme import Imobme
from patrimar_dependencies.functions import P

def prep_mp(grupos):
    imob = Imobme(
        login="renan.oliveira@patrimar.com.br",
        password="`+&tQ0Gu7@uJx1",
        url="https://patrimarengenharia.imobme.com",
        headless=False,
        debug=False
    )

    try:
        resultado = imob.extrair_relatorios(grupos)
    except Exception as err:
        print(P((type(err), err), color='red'))
        resultado = []
    finally:
        imob._encerrar()
        del imob

    return resultado    
    
if __name__ == "__main__":
    lista_relatorios = [
        [
            'imobme_empreendimento',
        #     'imobme_controle_vendas',
        #     'imobme_controle_vendas_90_dias',
        #     'imobme_contratos_rescindidos'
         ],
         [
             'imobme_contratos_rescindidos_90_dias',
        #     'imobme_dados_contrato',
        #     'imobme_previsao_receita',
        #     'imobme_relacao_clientes'
         ],
         [
             'imobme_relacao_clientes_x_clientes',
        #     'imobme_cadastro_datas',
        #     'recebimentos_compensados',
        #     'imobme_controle_estoque'
         ]
 ]

    num_process = mp.cpu_count() - 1 if len(lista_relatorios) > (mp.cpu_count() - 1) else len(lista_relatorios)
    
    pool = mp.Pool(processes=num_process)
    resultados = pool.map(prep_mp, lista_relatorios)
    pool.close()
    pool.join()
    
    lista_arquivos = []
    if resultados:
        for i, resultado in enumerate(resultados):
            for arquivo in resultado:
                lista_arquivos.append(arquivo)
    
    
    