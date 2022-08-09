from os import path
from sqlite3 import Cursor
import pandas as pd
from lib2to3.pgen2 import driver
from multiprocessing import connection
import pyodbc
import glob
import pyautogui
import os
from ConectarSqlServer import Conexao, retorna_conexao_sql
  

dados = glob.glob(Conexao.caminho)
for i in dados:
 tabela = pd.read_excel(i, sheet_name=None)
 for index in tabela: 
  nome_aba = index
  print(nome_aba)
  	     
  query = f"""CREATE TABLE {nome_aba}
    (DATA date NULL,
    CODIGO varchar(50) NULL,
    CENTRODECUSTO varchar(100) NULL,
    POSTODESERVICO varchar(500) NULL,
    FUNCIONARIO varchar(500) NULL,
    TIPO varchar(100) NULL,
    Refeicao varchar(50) NULL,
    Ausencia nvarchar(50) NULL,
    DataReposicao varchar(200) NULL,
    OBSINSPETOR varchar(600) NULL,
    OBSSUPERVISOR varchar(600) NULL,
    StatusSupervisor varchar(50) NULL,
    OBSCOORDENADOR varchar(600) NULL,
    StatusCoordenador varchar(50) NULL,
    )"""
 
 cursor = Conexao.retorna_conexao_sql()
 #executa query dentro do DB.
 cursor.execute(query)
 print("Query aqui: ", query)
 #utilizamos o commit quando vamos efetivar(comitar) o banco de dados (Criar, excluir ou atualizar).
 cursor.commit()
 
pyautogui.alert("Tabelas criadas com sucesso!\nProcesso finalizado.")







