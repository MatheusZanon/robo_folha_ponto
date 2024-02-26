"""COMPONENTE PARA CAIXA DE DIALOGO E INPUT DE USUARIO DO PYTHON"""
import tkinter as tk;
from datetime import date, timedelta

class DialogBox:
    def __init__(self, master):
        self.particao = None
        self.data1 = None
        self.data2 = None

        self.master = master
        self.master.title("Parâmetros Iniciais") # Título da Janela
        self.master.geometry("400x300")  # Tamanho da janela

        self.particao_var = tk.StringVar(self.master)
        self.data1_var = tk.StringVar(self.master)
        self.data2_var = tk.StringVar(self.master)

        self.error_message_label = tk.Label(master, text="", fg="red")
        self.error_message_label.pack()

        self.setup_ui()

    def setup_ui(self):
        # Título para o menu suspenso
        self.dropdown_title_label = tk.Label(self.master, text="Partição do Drive:")
        self.dropdown_title_label.pack()

        # Dropdown menu
        particoes = ["H", "I", "J", "K", "L", "M"]
        self.particao_var.set(particoes[0])  # Definir particao padrão
        menu_particoes = tk.OptionMenu(self.master, self.particao_var, *particoes)
        menu_particoes.pack()

        # Label e campo para data inicio
        label_data1_var = tk.Label(self.master, text="Data de Início:")
        label_data1_var.pack()
        entrada_data1_var = tk.Entry(self.master, textvariable=self.data1_var)
        entrada_data1_var.pack()

        # Label e campo para data fim
        label_data2_var = tk.Label(self.master, text="Data de Fim:")
        label_data2_var.pack()
        entrada_data2_var = tk.Entry(self.master, textvariable=self.data2_var)
        entrada_data2_var.pack()

        # Botão de envio
        botao_submit = tk.Button(self.master, text="Enviar", command=self.on_submit)
        botao_submit.pack()

    def on_submit(self):
        # Ação a ser realizada quando o botão Submit for pressionado
        self.particao = self.particao_var.get()
        self.data1 = self.data1_var.get()
        self.data2 = self.data2_var.get()

        self.master.destroy()
        


"""COLOCAR NO CODIGO PRINCIPAL
def main():
    root = tk.Tk()
    app = DialogBox(root)
    root.mainloop()
    return app.dropdown, app.mes, app.ano

if __name__ == "__main__":
    dropdown, mes, ano = main()
    print(f"Valor escolha: {dropdown}")
    print(f"Valor digitado 1: {mes}")
    print(f"Valor digitado 2: {ano}")
""" 