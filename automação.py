import pandas as pd
from datetime import datetime, timedelta
import time
import pyautogui as py
import pyperclip
import tkinter as tk
from tkinter import ttk, messagebox
import os, sys


# Função para encontrar arquivos

def get_caminho_arquivo(nome):
    if getattr(sys, 'frozen', False):
        # Caminho do executável
        base_path = os.path.dirname(sys.executable)
    else:
        # Caminho do script
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, nome)


planilha = get_caminho_arquivo("alimentaçãoplanilha.xlsx")
icon_arquivo = get_caminho_arquivo("icone.ico")



# Função principal de envio

def processar_envios():

    botao_iniciar.config(state="disabled")
    progresso["value"] = 0
    janela.update_idletasks()

    try:
        df = pd.read_excel(planilha)

        df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True, errors='coerce')

        hoje = datetime.now().date()

        amanha = hoje + timedelta(days=1)
        depois_amanha = hoje + timedelta(days=2)

        vencendo = df[
            (df['DATA'].dt.date == amanha) |
            (df['DATA'].dt.date == depois_amanha)
        ]

        if vencendo.empty:
            messagebox.showinfo("Aviso", "Nenhum cliente vencendo.")
            botao_iniciar.config(state="normal")
            return

        total = len(vencendo)
        passo = 100 / total

        # Abrir WhatsApp
        py.press('win')
        time.sleep(1)
        py.write('whatsapp')
        time.sleep(1)
        py.press('enter')
        time.sleep(5)  

        
        # Função de envio
        
        def enviar_mensagem(numero, mensagem):

            print("Enviando para:", numero)

            # Abrir busca
            py.hotkey('ctrl', 'f')
            time.sleep(1)

            # Limpar busca
            py.hotkey('ctrl', 'a')
            py.press('backspace')
            time.sleep(0.5)

            # Digitar número
            py.write(numero)
            time.sleep(2)

            # Entrar na conversa
            py.press('enter')
            time.sleep(3)

            # Enviar mensagem
            pyperclip.copy(mensagem)
            py.hotkey('ctrl', 'v')
            time.sleep(1)
            py.press('enter')
            time.sleep(2)

            # Voltar tela inicial
            py.hotkey('esc')
            time.sleep(1)

        # loop principal 

        for _, row in vencendo.iterrows():
            numero = str(row["CONTATOS"])
            nome = row["NOME"]
            data_venc = row["DATA"].strftime("%d/%m/%Y")

            mensagem = (
                f"Olá {nome}, espero que esteja tudo bem!\n"
                f"Passando pra avisar que seu plano vence no dia {data_venc}\n"
                f"Qualquer dúvida estamos por aqui!\n"
                f"PIX: 19989465723"
            )

            enviar_mensagem(numero, mensagem)

            progresso["value"] += passo
            janela.update_idletasks()

        messagebox.showinfo("Sucesso", "Mensagens enviadas!")

    except Exception as e:
        messagebox.showerror("Erro", str(e))

    botao_iniciar.config(state="normal")


# Interface Teste Gráfica

janela = tk.Tk()
janela.title("Automação de Mensagens")
janela.geometry("420x350")
janela.configure(bg="#1b1624")

if os.path.exists(icon_arquivo):
    janela.iconbitmap(icon_arquivo)

janela.resizable(False, False)

# Centralizar
janela.update_idletasks()
larg = janela.winfo_width()
alt = janela.winfo_height()
x = (janela.winfo_screenwidth() // 2) - (larg // 2)
y = (janela.winfo_screenheight() // 2) - (alt // 2)
janela.geometry(f"{larg}x{alt}+{x}+{y}")

# Título
titulo = tk.Label(
    janela, text="Automação",
    fg="#b56bff", bg="#1b1624",
    font=("Segoe UI", 24, "bold")
)
titulo.pack(pady=(30, 5))

sub = tk.Label(
    janela, text="Envio automático para clientes",
    fg="#a898be", bg="#1b1624",
    font=("Segoe UI", 12)
)
sub.pack(pady=(0, 25))

# Animação
def animar():
    cor = "#d188ff" if titulo.cget("fg") == "#b56bff" else "#b56bff"
    titulo.config(fg=cor)
    janela.after(600, animar)

animar()

# Botão
def on_enter(e):
    botao_iniciar.config(bg="#9b4dff")

def on_leave(e):
    botao_iniciar.config(bg="#7d2cff")

botao_iniciar = tk.Button(
    janela,
    text="INICIAR",
    font=("Segoe UI", 14, "bold"),
    bg="#7d2cff",
    fg="white",
    activebackground="#9b4dff",
    activeforeground="white",
    bd=0,
    width=16,
    height=2,
    cursor="hand2",
    command=processar_envios
)
botao_iniciar.pack(pady=10)

botao_iniciar.bind("<Enter>", on_enter)
botao_iniciar.bind("<Leave>", on_leave)

# Barra de progresso
style = ttk.Style()
style.theme_use("clam")
style.configure(
    "TProgressbar",
    troughcolor="#312d3c",
    bordercolor="#312d3c",
    background="#b56bff"
)

progresso = ttk.Progressbar(janela, length=260, style="TProgressbar")
progresso.pack(pady=25)

janela.mainloop()