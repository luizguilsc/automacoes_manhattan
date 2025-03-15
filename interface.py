from tkinter import *
from tkinter import filedialog, messagebox
import tkinter as tk
import pandas as pd
import os

class TelaPrincipal:
    def __init__(self, root):
        self.root = root
        self.root.title("Título da Janela")
        
        # Campo de E-mail
        self.label_email = tk.Label(self.root, text="E-mail").grid(column=0, row=0)
        self.entry_email = tk.Entry(self.root, width=40)
        self.entry_email.grid(column=0, row=1)
        
        # Campo de Senha
        self.label_senha = tk.Label(self.root, text="Senha").grid(column=0, row=2)
        self.entry_senha = tk.Entry(self.root, width=40, show="*")
        self.entry_senha.grid(column=0, row=3)  # Oculta a senha
        
        # Botão Salvar Login
        self.button_login = tk.Button(self.root, text="Salvar Login", command=self.salvando_login)
        self.button_login.grid(column=0, row=4)
        
        # Botão Importar Planilha
        self.importar_button = tk.Button(self.root, text="Importar", command=self.importar_planilha)
        self.importar_button.grid(column=2, row=3)
        
        # Botão Executar função do módulo automacao
        self.executar_button = tk.Button(self.root, text="Executar", command=self.executar_funcao)
        self.executar_button.grid(column=2, row=4)

        # Comentários e Reason code
        self.ref1_label = tk.Label(self.root, text= "Comentário 1").grid(column= 5, row= 1)
        self.ref1_entry = tk.Entry(self.root, width= 40)
        self.ref1_entry.grid(column= 5, row= 2)
        self.ref2_label = tk.Label(self.root, text= "Comentário 2").grid(column= 6, row= 1)
        self.ref2_entry = tk.Entry(self.root, width= 40)
        self.ref2_entry.grid(column= 6, row= 2)

        # Reason Code
        self.reasoncode_label = tk.Label(self.root, text="Selecione o Reason Code")
        self.reasoncode_label.grid(column=5, row=3)
        
        # Entry para exibir o Reason Code selecionado
        self.reasoncode_entry = tk.Entry(self.root, width=40)
        self.reasoncode_entry.grid(column=5, row=5)
        
        # Menu suspenso (OptionMenu)
        self.reason_codes = ["T3", "82", "72", "M1"]
        self.selected_reasoncode = StringVar(self.root)
        self.selected_reasoncode.set(self.reason_codes[0]) # Valor padrão
        
        self.reasoncode_menu = tk.OptionMenu(self.root, self.selected_reasoncode, *self.reason_codes, command=self.update_entry)
        self.reasoncode_menu.grid(column=5, row=4)
        
    def update_entry(self, value):
        self.reasoncode_entry.delete(0, END)
        self.reasoncode_entry.insert(0, value)

        
    def salvando_login(self):
        user, password = self.entry_email.get(), self.entry_senha.get()
        print(f"Email: {user} Senha: {password}")
        messagebox.showinfo("Login", "Login salvo com sucesso!")
    
    def importar_planilha(self):
        try:
            file = filedialog.askopenfilename(title="Selecione o Arquivo", filetypes=[("Excel Files", "*.csv;*.xlsm")])
            if file:
                if file.endswith(".csv"):
                    self.df = pd.read_csv(file)
                else:
                    self.df = pd.read_excel(file)

                messagebox.showinfo("Sucesso", f"Arquivo carregado: {os.path.basename(file)}")
                
                # Exibir uma prévia dos dados no console para verificar
                print("Primeiras linhas da planilha:")
                print(self.df.head())

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao importar planilha: {e}")

    def executar_funcao(self):
        try:
            from automacao import Automation, Login

            email = self.entry_email.get()
            senha = self.entry_senha.get()
            
            if not email or not senha:
                messagebox.showerror("Erro", "Preencha o e-mail e a senha antes de continuar.")
                return
            
            if not hasattr(self, 'df') or self.df.empty:
                messagebox.showerror("Erro", "Nenhuma planilha foi importada!")
                return

            automation = Automation()
            automation.setup_driver()

            login = Login(automation.driver, email, senha)
            login.logando()

            automation.selecionar_tela()
            automation.reason_code_auto(self.df)  # Passando os dados da planilha para a automação

            messagebox.showinfo("Sucesso", "Execução concluída com sucesso!")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar função: {e}")



class Lendo_Planilha():
    def __init__(self, root):
        self._root = root




if __name__ == "__main__":
    root = tk.Tk()
    app = TelaPrincipal(root)
    root.mainloop()