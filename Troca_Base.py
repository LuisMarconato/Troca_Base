#IMPORT
import logging
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
import sys
import time
from setuptools import Command
import win32service
import win32serviceutil

#1- ESTRUTURA DA JANELA:
janela = Tk()
janela.title("Gdoor - Sistemas")
janela["background"]= "orange red"

#1.1- Dimensões da janela
largura = 482
altura = 280

#1.2- RESOLUÇÃO DO NOSSO SISTEMA
largura_janela = janela.winfo_screenwidth()
altura_janela = janela.winfo_screenheight()

#1.3- POSIÇÃO DA JANELA
posx = largura_janela/2 - largura/2
posy = altura_janela/2 - altura/2

#1.4- DEFINIR A GEOMETRY
janela.geometry("%dx%d+%d+%d" % (largura,altura,posx,posy))

#1.5- DEFINIR ESPAÇO DO STATUS
LabelFrame  = LabelFrame(janela,text="STATUS:",font=("Arial",16,"bold"), background="orange red", labelanchor=N)
LabelFrame.pack(fill="both", expand="yes")

#1.6- TEXTO STATUS
texto_status = Label(LabelFrame,text="",font=("Arial",12),background="orange red")
texto_status.pack(anchor=CENTER, fill=Y, expand="yes")

#1.7- TEXTO EXECUÇÃO TROCA DE SENHA 
texto_senha = Label(LabelFrame, text="",font=("Arial",12),background="orange red")
texto_senha.pack(anchor=CENTER, fill=Y, expand="yes")

#1.8- TEXTO SELEÇÃO COMBOBOX
seleciona_modulo = Label(janela,text='SELECIONA O MÓDULO GDOOR:',font=("Arial",12,"bold"),background="orange red")
seleciona_modulo.pack( expand="yes")

#1.9- VARIAVEL GLOBAL DOS MÓDULOS
lista_modulos = ["PRO", "SLIM","MEI"]
caminho_pro='C:/GDOOR Sistemas/GDOOR PRO/'
caminho_slim='C:/GDOOR Sistemas/GDOOR SLIM/'
caminho_mei='C:/GDOOR Sistemas/GDOOR MEI/'

#1.10- CRIAÇÃO COMBOBOX
cb_modulos = ttk.Combobox(janela,values=lista_modulos)
cb_modulos.set(lista_modulos[0])
cb_modulos.pack(fill="both", expand="yes")
combo = 0

#2- VERIFICA SE O DIRETORIO JA ESTA CRIADO, SE NÃO TIVER SERÁ CRIADO
#GDOORPRO
if os.path.exists(f'{caminho_pro}'):
    if not os.path.exists(f'{caminho_pro}/Backup Tecnico')and not os.path.exists(f'{caminho_pro}/Base Cliente'):
            os.makedirs(f'{caminho_pro}/Base Cliente')
            os.makedirs(f'{caminho_pro}/Backup Tecnico')
#GDOORSLIM            
if os.path.exists(f'{caminho_slim}'):
    if not os.path.exists(f'{caminho_slim}/Backup Tecnico')and not os.path.exists(f'{caminho_slim}/Base Cliente'):
            os.makedirs(f'{caminho_slim}/Base Cliente')
            os.makedirs(f'{caminho_slim}/Backup Tecnico')
#GDOORMEI
if os.path.exists(f'{caminho_mei}'):
    if not os.path.exists(f'{caminho_mei}/Backup Tecnico')and not os.path.exists(f'{caminho_mei}/Base Cliente'):
            os.makedirs(f'{caminho_mei}/Base Cliente')
            os.makedirs(f'{caminho_mei}/Backup Tecnico')

#3-FUNÇÕES
#3.1 - PARAR/INICIAR O GUARD
def PARAR_GUARD():
    win32serviceutil.StopService("gdoorGuard_serv")
def INICIAR_GUARD():
    win32serviceutil.StartService("gdoorGuard_serv")

#3.2 - PARAR/INICIAR FIREBIRD    
def PARAR_FIREBIRD():
    win32serviceutil.StopService("FirebirdGuardianDefaultInstance")
def INICIAR_FIREBIRD():
    win32serviceutil.RestartService("FirebirdGuardianDefaultInstance") 

#3.3- VERIFICAÇÃO DOS SERVIÇOS( SE FOR =1 SERVIÇO ESTA PARADO SE FOR =4 SERVIÇO INCIADO)
def verifica_servicos_guard():
    verifica_guard= win32serviceutil.QueryServiceStatus('gdoorGuard_serv')[1]
    if verifica_guard ==1:
        return False
    elif verifica_guard ==4:
        return True
def verifica_servicos_firebird():
    verifica_firebird= win32serviceutil.QueryServiceStatus('FirebirdGuardianDefaultInstance')[1]
    if verifica_firebird ==1:
        return False
    if verifica_firebird ==4:
        return True

#3.4- FUNÇÃO DO BOTÃO TROCAR
def TROCA():
    #3.4.0- CONDICIONAL QUE ARMAZENA O CAMINHO DA BASE NA VARIAVEL modulo
    combo = cb_modulos.get()
    if combo == lista_modulos[0]:
        modulo = caminho_pro
    elif combo == lista_modulos[1]:
        modulo = caminho_slim
    elif combo == lista_modulos[2]:
        modulo = caminho_mei
    
    #3.4.1- VERIFICAR SE OS SERVICOS ESTÃO INICIADOS E PARA SERVIÇOS DO GUARD/FIREBIR
    if verifica_servicos_firebird()==True and verifica_servicos_guard()==True:
        PARAR_GUARD()
        PARAR_FIREBIRD()
        time.sleep(5)
    
    #3.4.2- VERIFICA SE OS SERVIÇOS ESTÃO PARADOS PARA COMEÇAR A AÇÃO DE TROCA DE BASE
    if verifica_servicos_firebird()==False and verifica_servicos_guard()==False:
         #3.4.3- RENOMEAR O DATAGES QUE ESTA RODANDO NO SISTEMA
         os.rename(f'{modulo}/DATAGES.FDB',f'{modulo}/DATAGES_OLD.FDB')
         #3.4.4- BACKUP DO DATAGES QUE ESTA RODANDO NO SISTEMA 
         shutil.move(f'{modulo}/DATAGES_OLD.FDB', f'{modulo}/Backup Tecnico/DATAGES_OLD.FDB')
         #3.4.5- COPIA TUDO DA PASTA BASE CLIENTE PARA A PASTA PASTA GDOOR.
         ori = f'{modulo}Base Cliente/' 
         des = modulo
         lista = os.listdir(ori) #lista separando apenas os arquivos do caminho.
         lista_len = len(lista) #ver quantos arquivos tem na pasta
         x=0
         #3.4.6-LAÇO PARA COPIAR OS ARQUIVOS PARA A PASTA DE DESTINO
         while x < lista_len :
            caminho_completo_ori = ori + lista[x] 
            caminho_completo_des = des + lista[x]
            shutil.move(caminho_completo_ori, caminho_completo_des)
            x+=1
         time.sleep(5) 

    #3.4.3- INICIAR O FIREBIRD PARA CONSEGUIR RODAR OS COMANDOS SQL
    if verifica_servicos_firebird()==False:
        INICIAR_FIREBIRD() 
    
    #3.4.4- VERIFICA SE O FIREBIRD ESTA INICIADO PARA CONSEGUIR CONECTAR NO DATAGES
    if verifica_servicos_firebird()==True:
        #3.4.5.1- APAGAR VERSÃO E TROCAR A SENHA DA BASE DO CLIENTE PARA 1
        conecta = firebirdsql.connect(user="SYSDBA", password="masterkey",database=f"{modulo}/DATAGES.FDB",host="localhost")
        cur = conecta.cursor()
        cur.execute("delete from versao_exe")
        cur.execute("update usuarios set senha = '1' where usuario = 'ADMINISTRADOR'")
        conecta.commit()
        #3.4.5.2- INICIAR O GUARD PARA CONSEGUIR ABRIR O GDOOR 
        INICIAR_GUARD()      
        
        #3.4.5-APRESENTAR MENSAGEM INDICADO STATUS
        status = 'Troca de base executada com sucesso!'
        senha = 'Senha do ADM alterada para 1'
        texto_status["text"] = status
        texto_senha["text"] = senha

        #3.4.6- HABILITANDO O BOTÃO REVERTER E DESABILITANDO O BOTÃO TROCA
        TROCA_button.configure(state='disable')
        REVERTE_button.configure(state='normal')
        
#3.5- FUNÇÃO DO BOTÃO REVERTE
def REVERTER():
    #3.5.1- CONDICIONAL QUE ARMAZENA O CAMINHO DA BASE NA VARIAVEL modulo
    combo = cb_modulos.get()
    if combo == lista_modulos[0]:
        modulo = caminho_pro
    elif combo == lista_modulos[1]:
        modulo = caminho_slim
    elif combo == lista_modulos[2]:
        modulo = caminho_mei
    
    #3.5.2- VERIFICAR SE OS SERVICOES ESTÃO INICIADOS E PARA OS SERVICOS DO GUARD/FIREBIRD
    if verifica_servicos_firebird()==True and verifica_servicos_guard()==True:
        PARAR_GUARD()
        PARAR_FIREBIRD()
        time.sleep(5) 
    
    #3.5.3- VERIFICA SE OS SERVIÇOS ESTÃO PARADOS PARA COMEÇAR A TROCA DE BASE
    if verifica_servicos_firebird()==False and verifica_servicos_guard()==False:    
        #3.3.4-RENOMEAR O DATAGES QUE ESTA RODANDO NO SISTEMA
        os.rename(f'{modulo}/DATAGES.FDB',f'{modulo}/DATAGES_CLIENTE.FDB')
        #3.3.5-MOVE O DATAGES DA PASTA GDOOR PARA A PASTA BASE CLIENTE
        shutil.move(f'{modulo}/DATAGES_CLIENTE.FDB', f'{modulo}/Base Cliente/DATAGES_CLIENTE.FDB')
        #3.3.6-RENOMEAR O DATAGES QUE DO TECNICO QUE ESTA NA PASTA BACKUP TECNICO
        os.rename(f'{modulo}/Backup Tecnico/DATAGES_OLD.FDB',f'{modulo}/Backup Tecnico/DATAGES.FDB')
        #3.3.7-MOVE O DATAGES DO TECNICO PARA A PASTA DO GDOOR PRO
        shutil.move(f'{modulo}/Backup Tecnico/DATAGES.FDB', f'{modulo}/DATAGES.FDB')
        time.sleep(5) 

    #3.5.4- INICIAR O FIREBIRD PARA CONSEGUIR RODAR OS COMANDOS SQL 
    if verifica_servicos_firebird()==False:
        INICIAR_FIREBIRD()
    
    #3.5.5- APAGAR VERSÃO E TROCAR A SENHA DA BASE DO TECNICO PARA 1
    if verifica_servicos_guard()==False: 
        conecta = firebirdsql.connect(user="SYSDBA", password="masterkey",database=f"{modulo}/DATAGES.FDB",host="localhost")
        cur = conecta.cursor()
        cur.execute("delete from versao_exe")
        cur.execute("update usuarios set senha = '1' where usuario = 'ADMINISTRADOR'")
        #3.5.6- INICIAR O GUARD PARA CONSEGUIR ABRIR O GDOOR 
        INICIAR_GUARD()
                        
        #3.5.6- APRESENTAR MENSAGEM INDICADO STATUS
        status = 'Base do tecnico restaurada com sucesso!'
        senha = 'Senha do ADM alterada para 1'
        texto_status["text"] = status
        texto_senha["text"] = senha

        #3.5.7- HABILITANDO O BOTÃO TROCA E DESABILITANDO O BOTÃO REVERTER
        TROCA_button.configure(state='normal')
        REVERTE_button.configure(state='disable')
                
#3.6- CRIAÇAO DA FUNÇÃO DELETE_VERSAO
def DELETE_VERSAO():
    #3.6.1- CONDICIONAL QUE ARMAZENA O CAMINHO DA BASE NA VARIAVEL modulo
    combo = cb_modulos.get()
    if combo == lista_modulos[0]:
        modulo = caminho_pro
    elif combo == lista_modulos[1]:
        modulo = caminho_slim
    elif combo == lista_modulos[2]:
        modulo = caminho_mei

    #3.6.2- VERIFICA SE O GUARD ESTA INICIADO E PARA O GUARD
    # OBS: PARA EVITAR QUE OCORRA CONFLITO COM O GUARD USANDO O DATAGES
    if verifica_servicos_guard() == True:
        PARAR_GUARD()
        time.sleep(5)
    
    #3.6.3- VERIFICA SE O SERVIÇO DO FIREBIRD ESTA INICIADO PARA CONSEGUIR CONECTAR NO BANCO DE DADOS
    if verifica_servicos_guard() == False:    
        #3.3.3- APAGAR A VERSÃO 
        conecta = firebirdsql.connect(user="SYSDBA", password="masterkey",database=f"{modulo}/DATAGES.FDB",host="localhost")
        cur = conecta.cursor()
        cur.execute("delete from versao_exe")
        conecta.commit()
        
        #3.6.4- APRESENTAR MENSAGEM INDICADO STATUS
        status = 'Comando SQL executado com sucesso!'
        texto_status["text"] = status
        status_servico = 0
        time.sleep(5)
        
    #3.6.5- INICIAR O GUARD PARA CONSEGUIR ABRIR O GDOOR
    if verifica_servicos_guard()== False:
        INICIAR_GUARD()
 
#4- BOTÕES
#4.1- BOTÃO DELETE
DELETE_button = Button(text='DELETE VERSÃO',height=1,width=15, command=DELETE_VERSAO)
DELETE_button.pack(fill="both", expand="yes")

#4.2- BOTÃO TROCAR
TROCA_button = Button(text='TROCAR BASE',height=1,width=15, command=TROCA)
TROCA_button.pack(fill="both", expand="yes")

#4.3- BOTÃO REVERTER
REVERTE_button = Button(text='REVERTER BASE',height=1,width=15,command=REVERTER)
REVERTE_button.pack(fill="both", expand="yes")

#5- VALIDAÇÃO DO BOTÃO REVERTER (APRIMORAR VARREDURA PARA HABILITAR BOTÕES DE AÇÕES)
if combo == 0:
    if len(os.listdir(f'{caminho_pro}Backup Tecnico')) == 0:
        REVERTE_button.configure(state='disable')
    
    elif len (os.listdir(f'{caminho_slim}Backup Tecnico')) == 0:
        REVERTE_button.configure(state='disable')
    
    elif len (os.listdir(f'{caminho_mei}Backup Tecnico')) ==0:
        REVERTE_button.configure(state='disable')

janela.mainloop()

