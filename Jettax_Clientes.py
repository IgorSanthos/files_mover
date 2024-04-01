import pandas as pd
import shutil
from pathlib import Path
from tkinter import messagebox
from PyQt5.QtWidgets import QApplication, QMessageBox#              janela de aviso

# Caminho do arquivo Excel
caminho_excel = Path(r'W:\Shared With Me\Clientes\Fenix Assessoria\Planilhas\RELAÇÃO JETTAX.xlsx')

# Ler a planilha Excel
df = pd.read_excel(caminho_excel, sheet_name='Jettax')

# Iterar sobre as linhas da planilha
for index, row in df.iterrows():
    caminho_origem = Path(row['ORIGEM'])   # Coluna F: Caminho de origem
    caminho_destino = Path(row['DESTINO'])  # Coluna I: Caminho de destino
    caminho_nfts = Path(row['NFTS'])    # Coluna G: Caminho de origem
    caminho_guias = Path(row['GUIA'])  # Coluna H: Caminho de origem
    caminho_destino2 = Path(row['DESTINO2']) # Coluna J: Caminho de destino
try: 
        # Verificar se o diretório de origem existe
    if caminho_guias.is_dir():
        # Iterar sobre os arquivos na pasta de origem
        for arquivo in caminho_guias.glob('*'):
            if 'guias' in arquivo.name:
                caminho_destino_arquivo = caminho_destino2/arquivo.name
            else:
                # Se não contiver 'enviadas' nem 'recebidas', ignorar o arquivo
                continue

            # Mover o arquivo para o destino
            shutil.move(arquivo, caminho_destino_arquivo)
            msg = print(f"Arquivo {arquivo.name} movido de {caminho_guias} para {caminho_destino_arquivo}")
    else:
        messagebox.showinfo("ERRO", f"Diretório de origem {caminho_guias} não encontrado. Movimento não realizado.")

    # Verificar se o diretório de origem existe
    if caminho_nfts.is_dir():
        # Iterar sobre os arquivos na pasta de origem
        for arquivo in caminho_nfts.glob('*'):
            if 'nfstomados' in arquivo.name:
                caminho_destino_arquivo = caminho_destino /'Serviços Tomados'/arquivo.name
            else:
                # Se não contiver 'enviadas' nem 'recebidas', ignorar o arquivo
                continue

            # Mover o arquivo para o destino
            shutil.move(arquivo, caminho_destino_arquivo)
            msg = print(f"Arquivo {arquivo.name} movido de {caminho_nfts} para {caminho_destino_arquivo}")
    else:
        messagebox.showinfo("ERRO", f"Diretório de origem {caminho_nfts} não encontrado. Movimento não realizado.")      


    # Verificar se o diretório de origem existe 
    if caminho_origem.is_dir():
        # Iterar sobre os arquivos na pasta de origem
        for arquivo in caminho_origem.glob('*'):
            # Verificar se o arquivo contém 'enviadas' no nome
            if 'enviadas' in arquivo.name:
                # Montar o caminho de destino para a pasta "Prestados"
                caminho_destino_arquivo = caminho_destino /'Serviços Prestados'/arquivo.name
            # Verificar se o arquivo contém 'recebidas' no nome
            elif 'recebidas' in arquivo.name:
                # Montar o caminho de destino para a pasta "Tomados"
                caminho_destino_arquivo = caminho_destino /'Serviços Tomados'/arquivo.name
            else:
                # Se não contiver 'enviadas' nem 'recebidas', ignorar o arquivo
                continue

            # Mover o arquivo para o destino
            shutil.move(arquivo, caminho_destino_arquivo)
            msg = print(f"Arquivo {arquivo.name} movido de {caminho_origem} para {caminho_destino_arquivo}")
    else:
        messagebox.showinfo("ERRO", f"Diretório de origem {caminho_origem} não encontrado. Movimento não realizado.")     


        app = QApplication([])

        msgBox = QMessageBox()
        msgBox.setWindowTitle('Alerta')
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setText('Arquivos pasta Jettax enviados para pasta Cliente')

        # Exibe a caixa de mensagem
        msgBox.exec_()
except Exception as error:
    messagebox.showinfo("ERRO", "MENSAGEM DE ERRO:", error)