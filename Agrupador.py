import sys
import os
import pandas as pd
from glob import glob
import customtkinter as ctk
from tkinter import messagebox
from tkinter.filedialog import askdirectory
from PIL import Image


if getattr(sys, 'frozen', False):
    caminho_base = sys._MEIPASS 
else:
    caminho_base = os.path.abspath(".")

caminho_imagem = os.path.join(caminho_base, "agrupador.jpg")

def process_files():
    # Abre a tela de seleção de pasta
    folder_path = askdirectory(title="Selecione a pasta com os arquivos")
    if not folder_path:
        print("Nenhuma pasta selecionada. O programa será encerrado.")
        return

    # Encontrar todos os arquivos que correspondem ao padrão
    file_pattern = os.path.join(folder_path, "settlement_detail_report_batch_*_*.xlsx")
    files = sorted(glob(file_pattern), key=lambda x: int(x.split('_')[4]))

    if not files:
        print("Nenhum arquivo encontrado na pasta selecionada.")
        return

    seen_numbers = set()
    missing_files = []
    dataframes = []

    # Processar todos os arquivos encontrados
    for file in files:
        current_num = int(file.split('_')[4])
        print(f"Processando: {file}")

        # Ignorar arquivos duplicados
        if current_num not in seen_numbers:
            seen_numbers.add(current_num)
            try:
                # Carregar o DataFrame
                df = pd.read_excel(file)

                # Se for o primeiro arquivo, armazena o cabeçalho e ignora apenas a primeira linha
                if not dataframes:
                    header_df = df.columns
                    df = df.iloc[1:]  # Remove a linha de cabeçalho

                # Adiciona o DataFrame à lista
                dataframes.append(df)

                # Checa se o próximo número na sequência está presente
                if current_num + 1 not in [int(f.split('_')[4]) for f in files]:
                    missing_files.append(current_num + 1)

            except Exception as e:
                print(f"Erro ao processar {file}: {e}")

    # Exibir aviso sobre arquivos faltantes
    if missing_files:
        missing_files_str = ', '.join([f"settlement_detail_report_batch_{num}_20241022" for num in missing_files])
        messagebox.showwarning("Arquivos Ausentes", f"{missing_files_str} está(ão) ausente(s) da pasta, sugiro que baixe o(s) arquivo(s) e refaça novamente.")

    # Verificar se há DataFrames para concatenar
    if dataframes:

        combined_df = pd.concat(dataframes, ignore_index=True)


        combined_df.columns = header_df


        output_file = os.path.join(folder_path, "resultado_agrupado.xlsx") 
        combined_df.to_excel(output_file, index=False)

        print(f"Arquivo salvo como {output_file}")
    else:
        print("Nenhum dado foi carregado para salvar.")


def start_process():
    process_files()

# Tela
app = ctk.CTk()
app.title("Agrupador de Arquivos")
app.geometry("500x500")

background_image = ctk.CTkImage(Image.open(caminho_imagem), size=(500, 500))
background_label = ctk.CTkLabel(app, image=background_image, text="")
background_label.place(x=0, y=0, relwidth=1, relheight=1)

start_button = ctk.CTkButton(app, text="Iniciar", command=start_process, fg_color="blue", hover_color="darkblue")
start_button.place(relx=0.5, rely=0.9, anchor="center")

app.mainloop()
