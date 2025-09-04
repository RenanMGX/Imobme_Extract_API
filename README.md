# Imobme Extract

Automação para extração, organização e exportação de relatórios do Imobme, integrada ao BotCity Maestro. Os arquivos deste repositório trabalham juntos como um único script orquestrado, com ponto de entrada local (`main.py`) e ponto de entrada para orquestração (`bot.py`).

## Sumário
- Visão geral
- Fluxo de alto nível
- Arquitetura do projeto
- Pré-requisitos
- Instalação
- Configuração (Credenciais e Parâmetros)
- Como executar
- Saídas e estrutura de pastas
- Dicas de desenvolvimento
- Solução de problemas

## Visão geral

O robô autentica no Imobme, seleciona e baixa relatórios, normaliza arquivos (Excel → JSON/CSV/XLSX) e grava no diretório de destino informado. Opcionalmente, a execução pode ser orquestrada pelo BotCity Maestro, que controla parâmetros, credenciais e telemetria.

## Fluxo de alto nível

1. Entrada de parâmetros (local ou Maestro)
2. Login no Imobme
3. Seleção e extração dos relatórios solicitados
4. Conversão e salvamento (JSON/CSV/XLSX) via `Entities/arquivos.py`
5. Limpeza de temporários (downloads e pastas auxiliares)

## Arquitetura do projeto

- `bot.py`
  - Integração com BotCity Maestro (`BotMaestroSDK`).
  - Lê parâmetros da tarefa (ex.: `crd`, `lista_relatorios`, `destino`, `extension`, `headless`, `quantidade`).
  - Usa `ExecuteAPP` (de `main.py`) para rodar o fluxo.
  - Reporta status, métricas e screenshots ao Maestro; integra com IA de apoio a erros (`patrimar_dependencies.gemini_ia`).

- `main.py`
  - Define a classe `ExecuteAPP` que orquestra a extração e o pós-processamento.
  - Ponto de entrada local com exemplo de `lista_relatorios`.
  - Utiliza `SharePointFolders` e `Credential` (do pacote `patrimar_dependencies`) para obter credenciais quando executado localmente.

- `Entities/imobme.py`
  - Camada de automação web (Selenium via `patrimar_dependencies.navegador_chrome`).
  - Login, navegação, seleção de relatórios, gerenciamento de downloads e limpeza.
  - Lista de relatórios válidos em `valid_relatorios` (ex.: `imobme_empreendimento`, `imobme_previsao_receita`, etc.).

- `Entities/arquivos.py`
  - Leitura/normalização de planilhas Excel com `xlwings` e `pandas`.
  - Funções para salvar em `JSON`, `CSV` ou `XLSX` respeitando nome/padrões.

- `Entities/exeptions.py`
  - Exceções específicas (ex.: `UrlError`, `RelatorioNotFound`, `BrowserClosed`).

- Outras pastas/arquivos
  - `downloads/`: pasta padrão de downloads temporários.
  - `resources/`: recursos (ex.: imagens).
  - `json/paths_register.json`: registro auxiliar de caminhos (se aplicável ao seu fluxo).
  - `build.bat`, `build.ps1`, `build.sh`: scripts de build/empacotamento.

## Pré-requisitos

- Windows 10/11 (execução validada em Windows)
- Python 3.10 a 3.12
- Google Chrome instalado
- Microsoft Excel (para `xlwings`) instalado (somente se for ler/transformar arquivos Excel)
- BotCity Maestro (opcional, apenas para orquestração)

## Instalação

1. Crie um ambiente virtual e instale as dependências:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install --upgrade -r requirements.txt
```

2. Verifique se o Chrome está instalado e atualizado.

## Configuração

### Credenciais

- Maestro: crie uma Credencial com o rótulo indicado em `crd` contendo chaves `login`, `password` e `url` do Imobme.
- Execução local: o exemplo em `main.py` mostra como carregar credenciais via `patrimar_dependencies.credenciais.Credential` usando `SharePointFolders`.

### Parâmetros principais

- `crd` (somente Maestro): label da credencial no Maestro.
- `lista_relatorios` (JSON): dicionário com os relatórios a extrair e como salvar.
- `destino`: caminho padrão (fallback) onde salvar os arquivos quando não informado por item.
- `extension`: formato padrão de saída (`JSON`, `CSV` ou `XLSX`) quando não informado por item.
- `headless`: `true|false` para execução do navegador sem interface.
- `quantidade`: tamanho do lote/grupo de relatórios por iteração (controle de paralelismo/lotes internos).

#### Exemplo de `lista_relatorios`

```json
{
  "imobme_empreendimento": {
    "destino": "C:/Users/usuario/Downloads/preparativo",
    "file_name": "DEFAULT",
    "extension": "JSON"
  },
  "imobme_relacao_clientes_x_clientes": {
    "destino": "C:/Users/usuario/Downloads/preparativo",
    "file_name": "DEFAULT",
    "extension": "JSON"
  },
  "imobme_previsao_receita": {
    "destino": "C:/Users/usuario/Downloads/preparativo",
    "file_name": "DEFAULT",
    "extension": "JSON"
  }
}
```

Campos por item:
- `destino`: diretório final do arquivo convertido.
- `file_name`: nome do arquivo; se `"DEFAULT"`, o nome é derivado do Excel original.
- `extension`: formato de saída (`JSON`, `CSV`, `XLSX`).
- `path_file` (opcional): caminho para um arquivo de origem específico quando aplicável.

## Como executar

### Local (desenvolvimento)

1. Ajuste o bloco `if __name__ == "__main__":` em `main.py` conforme suas credenciais e `lista_relatorios`.
2. Execute:

```powershell
python .\main.py
```

### Orquestrado (BotCity Maestro)

1. Publique o projeto no Maestro (ou utilize os scripts de build fornecidos).
2. Crie uma Tarefa e informe os parâmetros:
   - `task_name`: nome amigável da execução.
   - `crd`: label da credencial (com `login`, `password`, `url`).
   - `lista_relatorios`: JSON conforme exemplo.
   - `destino`, `extension`, `headless`, `quantidade` conforme sua necessidade.
3. Inicie a tarefa. O `bot.py` cuidará do restante e reportará métricas ao Maestro.

## Saídas e estrutura de pastas

- Arquivos finais: salvos no `destino` definido por item ou parâmetro padrão.
- Temporários: `downloads/` e uma pasta auxiliar de API criada em tempo de execução (limpas ao final quando possível).
- Nomes de arquivo: quando `file_name` for `DEFAULT`, o nome é baseado no nome da planilha original.

## Dicas de desenvolvimento

- Para mexer no fluxo web, veja `Entities/imobme.py` (login, navegação, seleção de relatórios, downloads).
- Para normalização/exportação, veja `Entities/arquivos.py` (funções `save_json_to`, `save_csv_to`, `save_excel_to`).
- Scripts de build: `build.ps1`/`build.bat`/`build.sh` (use conforme seu ambiente).

## Solução de problemas

- `ModuleNotFoundError: No module named 'botcity'`
  - Verifique se está usando o mesmo Python do ambiente virtual (ative-o) e reinstale `requirements.txt`.
- Chrome/Driver
  - Mantenha o Chrome atualizado. O wrapper de navegador em `patrimar_dependencies` gerencia o driver, mas versões muito antigas podem falhar.
- Excel em uso
  - O `xlwings` requer o Excel instalado. Feche instâncias do Excel; o código tenta encerrar/soltar locks antes de ler.
- Permissão de pastas
  - Garanta permissão de escrita no `destino` e em `downloads/`.
- Maestro: credenciais/labels
  - Confirme o `crd` e as chaves `login`, `password`, `url` na credencial.

---

Sinta-se à vontade para adaptar nomes de relatórios, formatos e destinos. Esta base foi desenhada para ser extensível e operar tanto localmente quanto orquestrada no Maestro.
