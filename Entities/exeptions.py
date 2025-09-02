class UrlError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
        
class RelatorioNotFound(Exception):
    def __init__(self, relat:str, *args: object) -> None:
        super().__init__(*args)
        
class BrowserClosed(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)