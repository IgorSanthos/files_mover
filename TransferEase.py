import pandas as pd
import shutil
from pathlib import Path
from PyQt5.QtWidgets import QApplication, QMessageBox

# Instância do QApplication
app = QApplication([])

# Caminho do arquivo Excel
caminho_excel = Path(r'W:\Shared With Me\Clientes\Fenix Assessoria\Planilhas\RELAÇÃO JETTAX.xlsx')

try:
    # Ler a planilha Excel
    df = pd.read_excel(caminho_excel, sheet_name='Jettax')

    # Iterar sobre as linhas da planilha
    for index, row in df.iterrows():
        caminho_origem = Path(row['ORIGEM'])   # Coluna F: Caminho de origem
        caminho_destino = Path(row['DESTINO'])  # Coluna I: Caminho de destino
        caminho_nfts = Path(row['NFTS'])    # Coluna G: Caminho de origem
        caminho_guias = Path(row['GUIA'])  # Coluna H: Caminho de origem
        caminho_destino2 = Path(row['DESTINO2']) # Coluna J: Caminho de destino

        # Movimentar arquivos de guias
        try:
            if caminho_guias.is_dir():
                for arquivo in caminho_guias.glob('*'):
                    caminho_destino_arquivo = caminho_destino2 / arquivo.name
                    shutil.move(arquivo, caminho_destino_arquivo)
                    print(f"Arquivo {arquivo.name} movido de {caminho_guias} para {caminho_destino_arquivo}")
            else:
                print(None, "ERRO", f"Diretório de origem {caminho_guias} não encontrado. Movimento não realizado.")
        except Exception as error:
            print(None, "ERRO", f"Ocorreu um erro ao mover os arquivos de guias: {error}")

        # Movimentar arquivos de notas fiscais NFTS
        try:
            if caminho_nfts.is_dir():
                for arquivo in caminho_nfts.glob('*'):
                    if 'nfs' in arquivo.name:
                        caminho_destino_arquivo = caminho_destino / 'Serviços Tomados' / arquivo.name
                        shutil.move(arquivo, caminho_destino_arquivo)
                        print(f"Arquivo {arquivo.name} movido de {caminho_nfts} para {caminho_destino_arquivo}")
            else:
                print(None, "ERRO", f"Diretório de origem {caminho_nfts} não encontrado. Movimento não realizado.")
        except Exception as error:
            print(None, "ERRO", f"Ocorreu um erro ao mover os arquivos de notas fiscais: {error}")

        # Movimentar arquivos de Prestados Tomados
        try:
            if caminho_origem.is_dir():
                for arquivo in caminho_origem.glob('*'):
                    if 'enviadas' in arquivo.name:
                        caminho_destino_arquivo = caminho_destino / 'Serviços Prestados' / arquivo.name
                    elif 'recebidas' in arquivo.name:
                        caminho_destino_arquivo = caminho_destino / 'Serviços Tomados' / arquivo.name
                    else:
                        continue
                    shutil.move(arquivo, caminho_destino_arquivo)
                    print(f"Arquivo {arquivo.name} movido de {caminho_origem} para {caminho_destino_arquivo}")
            else:
                QMessageBox.warning(None, "ERRO", f"Diretório de origem {caminho_origem} não encontrado. Movimento não realizado.")
        except Exception as error:
            QMessageBox.critical(None, "ERRO", f"Ocorreu um erro ao mover os arquivos de origem: {error}")

    # Exibir mensagem informando que os arquivos foram movidos com sucesso
    QMessageBox.information(None, "Sucesso", "Arquivos movidos com sucesso!")

except Exception as error:
    QMessageBox.critical(None, "ERRO", f"Ocorreu um erro: {error}")
    print(None, "ERRO", f"Ocorreu um erro: {error}")

# Executar o loop de eventos do PyQt
app.exec_()
