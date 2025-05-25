from tkinter import Tk, Text, Scrollbar, END, RIGHT, Y, LEFT, BOTH
import threading
import subprocess

def rodar_script():
    processo = subprocess.Popen(
        ["python", "app.py"],  # Certifique-se que app.py está no mesmo diretório
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        universal_newlines=True
    )
    for linha in processo.stdout:
        texto.insert(END, linha)
        texto.see(END)

# Interface
janela = Tk()
janela.title("Monitor de Execução do Manusis")
janela.geometry("800x600")

scrollbar = Scrollbar(janela)
scrollbar.pack(side=RIGHT, fill=Y)

texto = Text(janela, wrap='word', yscrollcommand=scrollbar.set)
texto.pack(side=LEFT, fill=BOTH, expand=True)

scrollbar.config(command=texto.yview)

threading.Thread(target=rodar_script, daemon=True).start()

janela.mainloop()