#criando classe
from msilib.schema import ServiceControl
from sqlite3 import Cursor
from sre_parse import State
from textwrap import fill
from tkinter import *
from tkinter import ttk
from turtle import width
from firebirdsql import *
import shutil
import os as os
import subprocess
import sys
import time
from setuptools import Command
import wmi
import win32service
import win32serviceutil

#1-ESTRUTURA DA JANELA:
janela = Tk()
janela.title("Gdoor - Sistemas")
#janela.wm_iconbitmap('troca.ico')
janela["background"]= "orange red"

#0-Dimensões da janela
largura = 482
altura = 280

#0.1-RESOLUÇÃO DO NOSSO SISTEMA
largura_janela = janela.winfo_screenwidth()
altura_janela = janela.winfo_screenheight()

#0.2- POSIÇÃO DA JANELA
posx = largura_janela/2 - largura/2
posy = altura_janela/2 - altura/2

#0.3- DEFINIR A GEOMETRY
janela.geometry("%dx%d+%d+%d" % (largura,altura,posx,posy))

#0.4- DEFINIR ESPAÇO DO STATUS
LabelFrame  = LabelFrame(janela,text="STATUS:",font=("Arial",16,"bold"), background="orange red", labelanchor=N)
LabelFrame.pack(fill="both", expand="yes")

#1.1-TEXTO STATUS
texto_status = Label(LabelFrame,text="",font=("Arial",12),background="orange red")
texto_status.pack(anchor=CENTER, fill=Y, expand="yes")

#1.2-TEXTO EXECUÇÃO TROCA DE SENHA 
texto_senha = Label(LabelFrame, text="",font=("Arial",12),background="orange red")
texto_senha.pack(anchor=CENTER, fill=Y, expand="yes")

#1.3-TEXTO SELEÇÃO COMBOBOX
seleciona_modulo = Label(janela,text='SELECIONA O MÓDULO GDOOR:',font=("Arial",12,"bold"),background="orange red")
#seleciona_modulo.place(x=0,y=118)
seleciona_modulo.pack( expand="yes")

#1.4-VARIAVEL GLOBAL DOS MÓDULOS
lista_modulos = ["PRO", "SLIM","MEI"]
caminho_pro='C:/GDOOR Sistemas/GDOOR PRO/'
caminho_slim='C:/GDOOR Sistemas/GDOOR SLIM/'
caminho_mei='C:/GDOOR Sistemas/GDOOR MEI/'

#1.5-CRIAÇÃO COMBOBOX
cb_modulos = ttk.Combobox(janela,values=lista_modulos)
cb_modulos.set(lista_modulos[0])
#cb_modulos.place(x=260,y=120,width=220)
cb_modulos.pack(fill="both", expand="yes")
combo = 0

#2-FUNÇÕES:
#2.1-VERIFICA SE O DIRETORIO JA ESTA CRIADO, SE NÃO TIVER SERÁ CRIADO
if os.path.exists(f'{caminho_pro}'):
    if not os.path.exists(f'{caminho_pro}/Backup Tecnico')and not os.path.exists(f'{caminho_pro}/Base Cliente'):
            os.makedirs(f'{caminho_pro}/Base Cliente')
            os.makedirs(f'{caminho_pro}/Backup Tecnico')
if os.path.exists(f'{caminho_slim}'):
    if not os.path.exists(f'{caminho_slim}/Backup Tecnico')and not os.path.exists(f'{caminho_slim}/Base Cliente'):
            os.makedirs(f'{caminho_slim}/Base Cliente')
            os.makedirs(f'{caminho_slim}/Backup Tecnico')
if os.path.exists(f'{caminho_mei}'):
    if not os.path.exists(f'{caminho_mei}/Backup Tecnico')and not os.path.exists(f'{caminho_mei}/Base Cliente'):
            os.makedirs(f'{caminho_mei}/Base Cliente')
            os.makedirs(f'{caminho_mei}/Backup Tecnico')

win32service.SERVICE_CONTROL_CONTINUE

#3-FUNÇÕES

#3.0-PARAR/INICIAR O GUARD E FIREBIRD
#3.0.1 - PARAR O GUARD
def PARAR_GUARD():
    win32serviceutil.StopService("gdoorGuard_serv")
def INICIAR_GUARD():
    time.sleep(1)  
    win32serviceutil.StartService("gdoorGuard_serv")

#3.0.2 - PARARO FIREBIRD    

def PARAR_FIREBIRD():
    win32serviceutil.StopService("FirebirdGuardianDefaultInstance")
def INICIAR_FIREBIRD():
    time.sleep(1)  
    win32serviceutil.RestartService("FirebirdGuardianDefaultInstance") 
     
#3.1-CRIAÇÃO DA FUNÇÃO TROCAR
def TROCA():
    
    #3.1.1-CONDICIONAL QUE ARMAZENA O ENDEREÇO NA VARIAVEL modulo
    combo = cb_modulos.get()
    if combo == lista_modulos[0]:
        modulo = caminho_pro
    elif combo == lista_modulos[1]:
        modulo = caminho_slim
    elif combo == lista_modulos[2]:
        modulo = caminho_mei

    #3.1.2- PARAR O GUARD
    PARAR_GUARD()
    PARAR_FIREBIRD()
    time.sleep(2) 

    #3.2.4-RENOMEAR O DATAGES
    os.rename(f'{modulo}/DATAGES.FDB',f'{modulo}/DATAGES_OLD.FDB')
    
    #3.3.5-BACKUP DO DATAGES DO TECNICO
    shutil.move(f'{modulo}/DATAGES_OLD.FDB', f'{modulo}/Backup Tecnico/DATAGES_OLD.FDB')
      
    #3.3.6-COPIA O DATAGES DA PASTA BASE CLIENTE PARA A PASTA PASTA GDOOR.
    ori = f'{modulo}Base Cliente/' 
    des = modulo

    lista = os.listdir(ori) #lista separando apenas os arquivos do caminho.
    lista_len = len(lista) #ver quantos arquivos tem na pasta
    x=0
    
    #3.3.7-LAÇO PARA COPIAR OS ARQUIVOS PARA A PASTA DE DESTINO
    while x < lista_len:
        caminho_completo_ori = ori + lista[x] 
        caminho_completo_des = des + lista[x]
        shutil.move(caminho_completo_ori, caminho_completo_des)
        x+=1
       
    #3.3.11- INICIAR O GUARD
    INICIAR_GUARD()
    INICIAR_FIREBIRD()

    #3.3.8-RODAR O DELETE NA BASE
    conecta = firebirdsql.connect(user="SYSDBA", password="masterkey",database=f"{modulo}/DATAGES.FDB",host="localhost")
    cur = conecta.cursor()
    cur.execute("delete from versao_exe")
    cur.execute("update usuarios set senha = '1' where usuario = 'ADMINISTRADOR'")
    conecta.commit()
   
    #3.3.9-APRESENTAR MENSAGEM INDICADO STATUS
    status = 'Troca de base executada com sucesso!'
    senha = 'Senha do ADM alterada para 1'
    texto_status["text"] = status
    texto_senha["text"] = senha

    #3.3.10- DESATIVANDO O BOTÃO CONTRÁRIO
    TROCA_button.configure(state='disable')
    REVERTE_button.configure(state='normal')
    
    #RODAR O RETAGUARDA NO FINAL
    #os.system('"%s' % 'gdoorGuard.exe')

#3.2-CRIAÇÃO DA FUNÇÃO REVERTE
def REVERTER():
           
    #3.2.1-CONDICIONAL QUE ARMAZENA O ENDEREÇO NA VARIAVEL modulo
    combo = cb_modulos.get()
    if combo == lista_modulos[0]:
        modulo = caminho_pro
    elif combo == lista_modulos[1]:
        modulo = caminho_slim
    elif combo == lista_modulos[2]:
        modulo = caminho_mei

    #3.2.2- PARAR O GUARD
    win32service.SERVICE_CONTROL_CONTINUE
    PARAR_GUARD()
    PARAR_FIREBIRD()
    time.sleep(2) 

    #3.2.4-RENOMEAR O DATAGES
    os.rename(f'{modulo}/DATAGES.FDB',f'{modulo}/DATAGES_CLIENTE.FDB')
    
    #3.2.5-MOVE O DATAGES DA PASTA GDOOR PARA A PASTA BASE CLIENTE
    shutil.move(f'{modulo}/DATAGES_CLIENTE.FDB', f'{modulo}/Base Cliente/DATAGES_CLIENTE.FDB')
    
    #3.2.6-RENOMEAR O DATAGES
    os.rename(f'{modulo}/Backup Tecnico/DATAGES_OLD.FDB',f'{modulo}/Backup Tecnico/DATAGES.FDB')

    #3.2.7-MOVE O DATAGES DO TECNICO PARA A PASTA DO GDOOR PRO
    shutil.move(f'{modulo}/Backup Tecnico/DATAGES.FDB', f'{modulo}/DATAGES.FDB')

    #3.2.11- INICIAR O GUARD
    win32service.SERVICE_CONTROL_CONTINUE
    INICIAR_GUARD()
    INICIAR_FIREBIRD()

    #3.2.8-RODAR O DELETE NA BASE
    conecta = firebirdsql.connect(user="SYSDBA", password="masterkey",database=f"{modulo}/DATAGES.FDB",host="localhost")
    cur = conecta.cursor()
    cur.execute("delete from versao_exe")
    cur.execute("update usuarios set senha = '1' where usuario = 'ADMINISTRADOR'")
    conecta.commit()
            
    #3.2.9-APRESENTAR MENSAGEM INDICADO STATUS
    status = 'Base do tecnico restaurada com sucesso!'
    senha = 'Senha do ADM alterada para 1'
    texto_status["text"] = status
    texto_senha["text"] = senha

    #3.2.10- DESATIVANDO O BOTÃO CONTRÁRIO
    TROCA_button.configure(state='normal')
    REVERTE_button.configure(state='disable')

    #RODAR O RETAGUARDA NO FINAL
    #os.system('"%s' % 'gdoorGuard.exe')
    
#3.3-CRIAÇAO DA FUNÇÃO DELETE_VERSAO
def DELETE_VERSAO():
    
    #3.3.2-CONDICIONAL QUE ARMAZENA O ENDEREÇO NA VARIAVEL modulo
    combo = cb_modulos.get()
    if combo == lista_modulos[0]:
        modulo = caminho_pro
    elif combo == lista_modulos[1]:
        modulo = caminho_slim
    elif combo == lista_modulos[2]:
        modulo = caminho_mei
    
    #3.3.1- PARAR O GUARD
    PARAR_GUARD()
    time.sleep(2)     

    #3.3.3-RODAR O DELETE NA BASE
    conecta = firebirdsql.connect(user="SYSDBA", password="masterkey",database=f"{modulo}/DATAGES.FDB",host="localhost")
    cur = conecta.cursor()
    cur.execute("delete from versao_exe")
    conecta.commit()
    
    #3.3.4-APRESENTAR MENSAGEM INDICADO STATUS
    status = 'Comando SQL executado com sucesso!'
    texto_status["text"] = status

    #3.3.5-INICIAR O GUARD
    win32service.SERVICE_CONTROL_CONTINUE
    INICIAR_GUARD()
    
    #RODAR O RETAGUARDA NO FINAL
    #os.system('"%s' % 'gdoorGuard.exe')

#BOTÃO DELETE
DELETE_button = Button(text='DELETE VERSÃO',height=1,width=15, command=DELETE_VERSAO)
#DELETE_button.place(x=2,y=150,width=130)
DELETE_button.pack(fill="both", expand="yes")

#BOTÃO TROCAR
TROCA_button = Button(text='TROCAR BASE',height=1,width=15, command=TROCA)
#TROCA_button.place(x=180,y=150,width=130)
TROCA_button.pack(fill="both", expand="yes")

#BOTÃO REVERTER
REVERTE_button = Button(text='REVERTER BASE',height=1,width=15,command=REVERTER)
#REVERTE_button.place(x=350,y=150,width=130)
REVERTE_button.pack(fill="both", expand="yes")

#VALIDAÇÃO DO BOTÃO REVERTER
if combo == 0:
    if len(os.listdir(f'{caminho_pro}Backup Tecnico')) == 0:
        REVERTE_button.configure(state='disable')
    
    elif len (os.listdir(f'{caminho_slim}Backup Tecnico')) == 0:
        REVERTE_button.configure(state='disable')
    
    elif len (os.listdir(f'{caminho_mei}Backup Tecnico')) ==0:
        REVERTE_button.configure(state='disable')

janela.mainloop()

