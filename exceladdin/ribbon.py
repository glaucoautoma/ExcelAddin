from xlwings import Ribbon

@Ribbon("Banco de Dados")
class DBRibbon:
    def __init__(self):
        self.connection_string = None
    
    @Ribbon.button(label="Configurar Conexão", image="connection")
    def configurar_conexao(self):
        from .functions import configurar_conexao
        self.connection_string = configurar_conexao()
    
    @Ribbon.button(label="Puxar Dados", image="download")
    def puxar_dados(self):
        from .functions import puxar_dados
        if not self.connection_string:
            print("Configure a conexão primeiro!")
            return
        puxar_dados(self.connection_string)
    
    @Ribbon.button(label="Enviar Dados", image="upload")
    def enviar_dados(self):
        from .functions import enviar_dados
        if not self.connection_string:
            print("Configure a conexão primeiro!")
            return
        enviar_dados(self.connection_string)