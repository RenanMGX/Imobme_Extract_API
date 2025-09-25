class Processos:
    
    @property
    def processados(self) -> int:
        return self.__processados
    
    @property
    def falhas(self) -> int:
        result = self.total - self.processados
        return result if result >= 0 else 0
    
    def __init__(self, value:int) -> None:
        self.total:int = value
        self.__processados:int = 0
        
    def add_processado(self, value:int=1):
        for _ in range(value):
            if (self.processados + 1) <= self.total:
                self.__processados += 1
