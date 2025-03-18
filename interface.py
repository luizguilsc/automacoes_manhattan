from tkinter import *
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
import pandas as pd
import os

class TelaPrincipal:
    def __init__(self, root):
        self.root = root
        self.root.title("Automações Manhattan")
        self.root.geometry("600x600")
        
        # Frame para os campos de login
        self.frame_login = ttk.LabelFrame(self.root, text="Login")
        self.frame_login.pack(padx=10, pady=10, fill=X)
        
        # Campo de E-mail
        self.label_email = ttk.Label(self.frame_login, text="E-mail")
        self.label_email.grid(column=0, row=0, sticky='w', padx=5, pady=5)
        self.entry_email = ttk.Entry(self.frame_login, width=40)
        self.entry_email.grid(column=1, row=0, padx=5, pady=5)
        
        # Campo de Senha
        self.label_senha = ttk.Label(self.frame_login, text="Senha")
        self.label_senha.grid(column=0, row=1, sticky='w', padx=5, pady=5)
        self.entry_senha = ttk.Entry(self.frame_login, width=40, show="*")
        self.entry_senha.grid(column=1, row=1, padx=5, pady=5)
        
        # Botão Salvar Login
        self.button_login = ttk.Button(self.frame_login, text="Salvar Login", command=self.salvando_login)
        self.button_login.grid(column=1, row=2, pady=10, sticky='e')
        
        # Frame para os botões de Importar e Executar
        self.frame_buttons = ttk.LabelFrame(self.root, text="Ações")
        self.frame_buttons.pack(padx=10, pady=10, fill=X)
        
        # Botão Importar Planilha
        self.importar_button = ttk.Button(self.frame_buttons, text="Importar", command=self.importar_planilha)
        self.importar_button.grid(column=0, row=0, padx=5, pady=5)
        
        # Label para o botão Verify
        self.verify_label = ttk.Label(self.frame_buttons, text="Verificação:")
        self.verify_label.grid(column=1, row=0, padx=5, pady=5)
        
        # Botão Executar Verify
        self.verify_button = ttk.Button(self.frame_buttons, text="Executar Verify", command=self.executar_verify)
        self.verify_button.grid(column=2, row=0, padx=5, pady=5)
        
        # Frame para os comentários e Reason Code
        self.frame_comentarios = ttk.LabelFrame(self.root, text="Comentários e Reason Code")
        self.frame_comentarios.pack(padx=10, pady=10, fill=X)
        
        # Comentários
        self.ref1_label = ttk.Label(self.frame_comentarios, text="Selecionar Comentário - 1:")
        self.ref1_label.grid(column=0, row=0, sticky='w', padx=5, pady=5)
        self.ref1_comentario_lista = ['Ilpn C/Bloqueio (82/72)', 'M1 - Origem 0014 P/ InventoryType P/ 1401', 'Trocando Status BOA P/ QEB', 'Trocando Status QEB P/ BOA', 'D15 - Débito 20%', 'FA - Débito Extravio 100%', 'DT - Débito Total 100%']
        self.comentarios = StringVar()
        self.coment_1 = ttk.Combobox(self.frame_comentarios, width=40, textvariable=self.comentarios)
        self.coment_1['values'] = self.ref1_comentario_lista
        self.coment_1.grid(column=1, row=0, padx=5, pady=5)
        
        self.ref2_label = ttk.Label(self.frame_comentarios, text="Digitar Comentário/Observação - 2: ")
        self.ref2_label.grid(column=0, row=1, sticky='w', padx=5, pady=5)
        self.ref2_entry = ttk.Entry(self.frame_comentarios, width=43)
        self.ref2_entry.grid(column=1, row=1, padx=5, pady=5)
        
        # Frame para Reason Code, Filial e Status
        self.frame_reason_filial_status = ttk.Frame(self.frame_comentarios)
        self.frame_reason_filial_status.grid(column=0, row=2, columnspan=2, pady=10, sticky='w')
        
        # Reason Code
        self.reasoncode_label = ttk.Label(self.frame_reason_filial_status, text="Reason Code")
        self.reasoncode_label.grid(column=0, row=0, padx=5)
        self.lista_reasoncode = ['T3', 'M1', '72', '82', 'FA', 'DT', 'AV', 'VF', 'VR']
        self.rc = StringVar()
        self.lista_itens = ttk.Combobox(self.frame_reason_filial_status, width=5, textvariable=self.rc)
        self.lista_itens['values'] = self.lista_reasoncode
        self.lista_itens.grid(column=1, row=0, padx=5)
        
        # Filial
        self.filial_field = ttk.Label(self.frame_reason_filial_status, text="Filial")
        self.filial_field.grid(column=2, row=0, padx=5)
        self.lista_filial = ["1401", "0014"]
        self.filial = StringVar()
        self.lista_filial_combobox = ttk.Combobox(self.frame_reason_filial_status, width=5, textvariable=self.filial)
        self.lista_filial_combobox['values'] = self.lista_filial
        self.lista_filial_combobox.grid(column=3, row=0, padx=5)
        
        # Status
        self.status_field = ttk.Label(self.frame_reason_filial_status, text="Status BOA/QEB")
        self.status_field.grid(column=4, row=0, padx=5)
        self.lista_status = ["BOA", "QEB"]
        self.status = StringVar()
        self.lista_status_combobox = ttk.Combobox(self.frame_reason_filial_status, width=5, textvariable=self.status)
        self.lista_status_combobox['values'] = self.lista_status
        self.lista_status_combobox.grid(column=5, row=0, padx=5)
        
        # Botão Executar função do módulo automacao
        self.executar_button = ttk.Button(self.frame_reason_filial_status, text="Executar Reason Code", command=self.executar_funcao)
        self.executar_button.grid(column=6, row=0, padx=5)
        
        # Display da Planilha
        self.display_planilha = ttk.LabelFrame(self.root, text='Planilha')
        self.display_planilha.pack(padx=10, pady=10, fill=BOTH, expand=True)
        self.data_display = Text(self.display_planilha, wrap=NONE)
        self.data_display.pack(padx=5, pady=5, fill=BOTH, expand=True)
                
    def update_entry(self, value):
        self.reasoncode_entry.delete(0, END)
        self.reasoncode_entry.insert(0, value)

    def salvando_login(self):
        user, password = self.entry_email.get(), self.entry_senha.get()
        # print(f"Email: {user} Senha: {password}")
        if user and password:
            print(f"Email: {user} Senha: {password}")
            messagebox.showinfo("Login", "Login salvo com sucesso!")
        else: 
            messagebox.showerror("Falha Login", "Preencher Login antes de continuar")

    def importar_planilha(self):
        try:
            file = filedialog.askopenfilename(title="Selecione o Arquivo", filetypes=[("Excel Files", "*.csv;*.xlsm")])
            if file:
                if file.endswith(".csv"):
                    self.df = pd.read_csv(file)
                else:
                    self.df = pd.read_excel(file)
                messagebox.showinfo("Sucesso", f"Arquivo carregado: {os.path.basename(file)}")
                
                # Exibir os dados da planilha no Text widget
                self.exibir_dados_planilha()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao importar planilha: {e}")

    def exibir_dados_planilha(self):
        self.data_display.delete(1.0, END)  # Limpar o conteúdo atual
        self.data_display.insert(END, self.df.to_string(index=False))  # Inserir os dados da planilha

    def executar_funcao(self):
        try:
            from automacao import Automation, Login
            email = self.entry_email.get()
            senha = self.entry_senha.get()
            
            # Verificação corrigida
            if not email or not senha:
                messagebox.showerror("Erro", "Preencha o e-mail e a senha antes de continuar.")
                return
            
            if not hasattr(self, 'df') or self.df.empty:
                messagebox.showerror("Erro", "Nenhuma planilha foi importada!")
                return
            
            reason_code = self.rc.get()
            filial = self.filial.get()
            status = self.status.get()
            comentario1 = self.comentarios.get()
            comentario2 = self.ref2_entry.get()
            automation = Automation()
            automation.setup_driver()
            login = Login(automation.driver, email, senha)
            login.logando()
            automation.selecionar_tela_inventory_details()
            
            # Processar a planilha e gerar feedback
            self.df = automation.reason_code_auto(self.df, reason_code, status, filial, comentario1, comentario2)
            
            # Exibir os dados da planilha com feedback no Text widget
            self.exibir_dados_planilha()
            
            messagebox.showinfo("Sucesso", "Execução concluída com sucesso!")
            automation.close_driver()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar função: {e}")
            if automation.driver:
                automation.close_driver()
    
    def executar_verify(self):
        try:
            from automacao import Automation, Login
            email = self.entry_email.get()
            senha = self.entry_senha.get()
            
            # Verificação corrigida
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
            automation.selecionar_asn()
            
            # Processar a planilha e gerar feedback
            self.df = automation.auto_verify(self.df)
            
            # Exibir os dados da planilha com feedback no Text widget
            self.exibir_dados_planilha()
            
            messagebox.showinfo("Sucesso", "Execução concluída com sucesso!")
            automation.close_driver()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar função: {e}")
            if automation.driver:
                automation.close_driver()
    

if __name__ == "__main__":
    root = tk.Tk()
    app = TelaPrincipal(root)
    root.mainloop()