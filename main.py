import pandas as pd
from time import sleep
import os
import PySimpleGUI as sg
from threading import Thread



class Mbot():
    def __init__(self):
        self.programa()
    
    def faz_janela(self):
        os.system('clear')
        sg.theme('DarkBlue')
            
        layout = [[sg.Text('Matrix Bot V0.5')],
                [sg.Text(' ', key=('-INFO-'))], 
                [sg.Button('Iniciar', key=('-INI-'), size=(35, 1))], 
                [sg.Button('Parar', key=('-PAR-'), size=(35, 1))],
                [sg.Text(' ')],
                [sg.Button('Sair', key=('SAI'), size=(35, 1))]]  
        
        return sg.Window('Principal', layout, location=(800,600), finalize=True, no_titlebar=True, grab_anywhere=True)

    def rodar(self):
        i = 0

        while True:
            info = self.janela['-INFO-']
            info.update(f"Rodou {i:.0f} vezes")

            i = i + 1
            # Define o arquivo
            file_name = 'Base.xlsx' 

            # Carrega o arquivo e ignora linhas inuteis
            df = pd.read_excel(file_name, skiprows=[0, 1])

            # Seleciona colunas e soma valores
            df = df.groupby(['Item Código','Descrição do item'], as_index=False)['Quantidade', 'Recebido'].sum()


            try:

                # Salva o arquivo
                df.to_excel(r'Final.xlsx', sheet_name='Receber', index = False)
                os.system('clear')
                #print("\n")
                info.update("Arquivo salvo com sucesso! ☺")
                #print("*" * 30, "\n")
                sleep(2)

            except Exception as e:
                print(f"\nErro ao salvar: {e}\n")

            if self.p == 0:
                info.update("Processo desligado!")
                break
            

    def programa(self):

        
        # Define a primeira janela a iniciar          
        window1 = self.faz_janela()    
            
                
        while True:                     
            
            janela, eventos, valores = sg.read_all_windows()     
            self.janela = janela
            

            if eventos == 'SAI' or eventos == 'Sair': #Eventos que fecham a janela
                janela.close()
                os.system('clear')
                break
            
            elif eventos == '-INI-': 
                janela['-INFO-'].update("Processo ligando...") 
                self.p = 1             
                processo = Thread(target=self.rodar)
                processo.start()   
                print(processo.is_alive())             

            elif eventos == '-PAR-':
                self.p = 0
                janela['-INFO-'].update("Processo desligando...")
                print(processo.is_alive())

                

if __name__=="__main__":
    Mbot()                                                 