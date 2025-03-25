import pandas as pd
from sqlalchemy import create_engine
import xlwings as xw

def configurar_conexao():
    """Configura a conexão com o banco de dados"""
    import tkinter as tk
    from tkinter import simpledialog
    
    root = tk.Tk()
    root.withdraw()
    
    server = simpledialog.askstring("Configuração", "Servidor:")
    database = simpledialog.askstring("Configuração", "Banco de dados:")
    username = simpledialog.askstring("Configuração", "Usuário:")
    password = simpledialog.askstring("Configuração", "Senha:", show='*')
    
    # SQL Server example - altere para seu banco de dados
    connection_string = (
        f"mssql+pyodbc://{username}:{password}@{server}/{database}"
        "?driver=ODBC+Driver+17+for+SQL+Server"
    )
    
    return connection_string

def puxar_dados(connection_string):
    """Puxa dados do banco para o Excel"""
    try:
        engine = create_engine(connection_string)
        
        # Exemplo: ler dados de uma tabela
        query = "SELECT * FROM SuaTabela"
        df = pd.read_sql(query, engine)
        
        # Escrever no Excel
        wb = xw.Book.caller()
        sheet = wb.sheets[0]
        sheet.clear()
        sheet.range('A1').value = df
        
        print("Dados puxados com sucesso!")
    except Exception as e:
        print(f"Erro ao puxar dados: {str(e)}")

def enviar_dados(connection_string):
    """Envia dados do Excel para o banco"""
    try:
        engine = create_engine(connection_string)
        wb = xw.Book.caller()
        sheet = wb.sheets[0]
        
        # Ler dados do Excel
        df = sheet.used_range.options(pd.DataFrame, index=False).value
        
        # Exemplo: enviar para uma tabela (substitua pelo seu código)
        df.to_sql('SuaTabela', engine, if_exists='append', index=False)
        
        print("Dados enviados com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar dados: {str(e)}")