from tkinter import *
from tkinter import filedialog, messagebox, ttk
import ttkbootstrap as ttk
import tkinter as tk
import pandas as pd
import os
import webbrowser

class TelaPrincipal:
    def __init__(self, root):
        self.root = root
        self.root.title("Automações Manhattan")
        self.root.geometry("800x600")

        # Create a Notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(padx=10, pady=10, fill=BOTH, expand=True)

        # Create tabs
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.tab3 = ttk.Frame(self.notebook)
        self.tab4 = ttk.Frame(self.notebook)

        # Add tabs to the notebook
        self.notebook.add(self.tab1, text="Login")
        self.notebook.add(self.tab2, text="Ações")
        self.notebook.add(self.tab3, text="Reason Code")
        self.notebook.add(self.tab4, text="Tema e Links")

        # Tab 1: Links e Tema
        self.create_tab1_content()

        # Tab 2: Login
        self.create_tab2_content()

        # Tab 3: Ações
        self.create_tab3_content()

        # Tab 4: Comentários e Reason Code
        self.create_tab4_content()

    def create_tab1_content(self):
        # Frame para links (GitHub, LinkedIn, Ajuda)
        self.frame_links = ttk.Frame(self.tab4)
        self.frame_links.pack(padx=5, pady=5, anchor="e")



        # Funções para abrir os links
        def abrir_github():
            webbrowser.open("https://github.com/luizguilsc")

        def abrir_linkedin():
            webbrowser.open("https://www.linkedin.com/in/luiz-guilherme-0b8a2a300/")

        def mostrar_ajuda():
            messagebox.showinfo("Sobre o Aplicativo", "Este é um software de automação para processos Manhattan.\n\nAntes de iniciar precisa preencher o campo de login e senha, salvando em seguida.\n\nPrimeira coisa a se fazer é extrair pelo export no Manhattan o Excel(.csv).\n\nPara etapas de EAD faturado importar apenas a planilha com as ASN, clicar em 'Receber EAD' - sempre conferir se foi efetuado com sucesso, em seguida clicar em 'Executar Verify'.\n\nPara realizar os Reason Codes, primeiro importar Excel(.csv) com a 'ILPN' e 'Item', preencher os comentário e certificar de selecionar o Reson Code correto, filial e Status.\n\nLembrando que o software apenas automotiza o que você faria manualmente, da mesma forma é necessária devia atênção quando selecionar a filial e status, caso execute com alguma informação errada, basta fechar o Chrome da automação que o processo será cancelado.\n\nNa parte inferior está o display da planilha importada, onde mostra o feedback quando as ASN ou ILPN são processadas")

        # Carregar ícones
        self.icon_github = PhotoImage(file="icons/github.png")
        self.icon_linkedin = PhotoImage(file="icons/linkedin.png")
        self.icon_ajuda = PhotoImage(file="icons/help.png")

        # Criando botões com ícones
        self.button_github = ttk.Button(self.frame_links, bootstyle="light-toolbutton", image=self.icon_github, command=abrir_github)
        self.button_github.grid(column=0, row=0, padx=5)
        
        self.button_linkedin = ttk.Button(self.frame_links, bootstyle="light-toolbutton", image=self.icon_linkedin, command=abrir_linkedin)
        self.button_linkedin.grid(column=1, row=0, padx=5)
        
        self.button_ajuda = ttk.Button(self.frame_links, bootstyle="primary-toolbutton", image=self.icon_ajuda, command=mostrar_ajuda)
        self.button_ajuda.grid(column=2, row=0, padx=5)


        # Frame principal para conter tema e links
        
        self.frame_principal = ttk.Frame(self.tab4)
        self.frame_principal.pack(padx=10, pady=10, fill=X)
        
        # Frame para seleção de tema
        self.frame_tema = ttk.LabelFrame(self.frame_principal, text="Tema do Aplicativo")
        self.frame_tema.grid(column=0, row=0, padx=10, pady=10, sticky='w')

        # Label para o ComboBox de temas
        self.label_tema = ttk.Label(self.frame_tema, text="Escolha o Tema:")
        self.label_tema.grid(column=0, row=0, padx=5, pady=5)

        # Variável para armazenar o tema selecionado
        self.tema_selecionado = tk.StringVar(value="darkly")  # Tema inicial
        # ComboBox para escolher o tema
        self.lista_tema = ["cosmo", "flatly", "journal", "litera", "lumen", "minty", "pulse", "sandstone", "united", "yeti", "morph", "darkly", "cyborg", "superhero", "solar"]
        
        self.combobox_tema = ttk.Combobox(self.frame_tema, values=self.lista_tema, textvariable=self.tema_selecionado, state="readonly")
        self.combobox_tema.grid(column=1, row=0, padx=5, pady=5)
        
        self.combobox_tema.set("litera")  # Define o tema inicial como "darkly"

        # Botão para aplicar o tema
        self.botao_tema = ttk.Button(self.frame_tema, text="Aplicar Tema", command=self.aplicar_tema)
        self.botao_tema.grid(column=2, row=0, padx=5, pady=5)

    def aplicar_tema(self):
        novo_tema = self.tema_selecionado.get()  # Obtém o tema selecionado
        self.root.style.theme_use(novo_tema)  # Aplica o novo tema

    def create_tab2_content(self):
         # Frame para os campos de login
         frame_login = ttk.LabelFrame(self.tab1, text='Login')
         frame_login.pack(padx=10, pady=10, fill=X)

         label_email = ttk.Label(frame_login, text='E-mail')
         label_email.grid(column=0, row=0, sticky='w', padx=5, pady=5)

         self.entry_email = ttk.Entry(frame_login, width=40)
         self.entry_email.grid(column=1, row=0, padx=5, pady=5)

         label_senha = ttk.Label(frame_login, text='Senha')
         label_senha.grid(column=0, row=1, sticky='w', padx=5, pady=5)

         self.entry_senha = ttk.Entry(frame_login, width=40, show='*')
         self.entry_senha.grid(column=1, row=1, padx=5, pady=5)

         button_login = ttk.Button(frame_login, text='Salvar Login', command=self.salvando_login)
         button_login.grid(column=1, row=2, pady=10, sticky='e')

    def create_tab3_content(self):
         frame_buttons = ttk.LabelFrame(self.tab2, text='Ações')
         frame_buttons.pack(padx=10, pady=10, fill=X)

         importar_button = ttk.Button(frame_buttons, text='Importar', command=self.importar_planilha)
         importar_button.grid(column=0, row=0, padx=5, pady=5)

         verify_label = ttk.Label(frame_buttons, text='Verificação:')
         verify_label.grid(column=3, row=0, padx=5, pady=5)

         verify_button = ttk.Button(frame_buttons, text='Executar Verify', command=self.executar_verify)
         verify_button.grid(column=4, row=0, padx=5, pady=5)

         ead_label = ttk.Label(frame_buttons, text='EAD Faturado:')
         ead_label.grid(column=1, row=0, padx=5, pady=5)

         ead_button = ttk.Button(frame_buttons, text='Receber EAD', command=self.executar_receiving_EAD)
         ead_button.grid(column=2, row=0, padx=5, pady=5)

    def create_tab4_content(self):
         frame_comentarios = ttk.LabelFrame(self.tab3, text='Comentários e Reason Code')
         frame_comentarios.pack(padx=10, pady=10, fill=X)

         ref1_label = ttk.Label(frame_comentarios, text='Selecionar Comentário - 1:')
         ref1_label.grid(column=0, row=0, sticky='w', padx=5, pady=5)

         ref1_comentario_lista = ['Ilpn C/Bloqueio (82/72)', 'M1 - Origem 0014 P/ InventoryType P/ 1401', 'Trocando Status BOA P/ QEB', 'Trocando Status QEB P/ BOA', 'D15 - Débito 20%', 'FA - Débito Extravio 100%', 'DT - Débito Total 100%']
         comentarios = StringVar()
         coment_1 = ttk.Combobox(frame_comentarios, width=40, textvariable=comentarios)
         coment_1['values'] = ref1_comentario_lista
         coment_1.grid(column=1, row=0, padx=5, pady=5)

         ref2_label = ttk.Label(frame_comentarios, text='Digitar Comentário/Observação - 2:')
         ref2_label.grid(column=0, row=1, sticky='w', padx=5, pady=5)

         ref2_entry = ttk.Entry(frame_comentarios, width=43)
         ref2_entry.grid(column=1, row=1, padx=5, pady=5)

         frame_reason_filial_status = ttk.Frame(frame_comentarios)
         frame_reason_filial_status.grid(column=0, row=2, columnspan=2, pady=10, sticky='w')

         reasoncode_label = ttk.Label(frame_reason_filial_status, text='Reason Code')
         reasoncode_label.grid(column=0, row=0, padx=5)

         lista_reasoncode = ['T3', 'M1', '72', '82', 'FA', 'DT', 'AV', 'VF', 'VR']
         rc = StringVar()
         lista_itens = ttk.Combobox(frame_reason_filial_status, width=5, textvariable=rc)
         lista_itens['values'] = lista_reasoncode
         lista_itens.grid(column=1, row=0, padx=5)

         filial_field = ttk.Label(frame_reason_filial_status, text='Filial')
         filial_field.grid(column=2, row=0, padx=5)

         lista_filial = ["1401", "0014"]
         filial = StringVar()
         lista_filial_combobox = ttk.Combobox(frame_reason_filial_status, width=5, textvariable=filial)
         lista_filial_combobox['values'] = lista_filial
         lista_filial_combobox.grid(column=3, row=0, padx=5)

         status_field = ttk.Label(frame_reason_filial_status, text='Status BOA/QEB')
         status_field.grid(column=4, row=0, padx=5)

         lista_status = ["BOA", "QEB"]
         status = StringVar()
         lista_status_combobox = ttk.Combobox(frame_reason_filial_status, width=5, textvariable=status)
         lista_status_combobox['values'] = lista_status
         lista_status_combobox.grid(column=5, row=0, padx=5)

         executar_button = ttk.Button(frame_reason_filial_status, text='Executar Reason Code', command=self.executar_funcao)
         executar_button.grid(column=6, row=0, padx=5)

         # Display da Planilha com Scroll
         display_planilha = ttk.LabelFrame(self.root, text='Planilha')
         display_planilha.pack(padx=10, pady=10, fill=BOTH, expand=True)

         scrollbar_x = ttk.Scrollbar(display_planilha, orient=HORIZONTAL)
         scrollbar_x.pack(side=BOTTOM, fill=X)

         scrollbar_y = ttk.Scrollbar(display_planilha, orient=VERTICAL)
         scrollbar_y.pack(side=RIGHT, fill=Y)

         self.data_display = Text(display_planilha, wrap=NONE, xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set)
         self.data_display.pack(padx=5, pady=5, fill=BOTH, expand=True)

         scrollbar_x.config(command=self.data_display.xview)
         scrollbar_y.config(command=self.data_display.yview)

    def update_entry(self, value):
        self.reasoncode_entry.delete(0, END)
        self.reasoncode_entry.insert(0, value)

    def salvando_login(self):
        user, password = self.entry_email.get(), self.entry_senha.get()
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
                self.exibir_dados_planilha()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao importar planilha: {e}")

    def exibir_dados_planilha(self):
        self.data_display.delete(1.0, END)
        self.data_display.insert(END, self.df.to_string(index=False))

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
            self.df = automation.reason_code_auto(self.df, reason_code, status, filial, comentario1, comentario2)
            self.exibir_dados_planilha()
            messagebox.showinfo("Sucesso", "Execução concluída com sucesso!")
            automation.close_driver()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar função: {e}")
            if 'automation' in locals() and automation.driver:
                automation.close_driver()

    def executar_verify(self):
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
            automation.selecionar_asn()
            self.df = automation.auto_verify(self.df)
            self.exibir_dados_planilha()
            messagebox.showinfo("Sucesso", "Execução concluída com sucesso!")
            automation.close_driver()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar função: {e}")
            if 'automation' in locals() and automation.driver:
                automation.close_driver()

    def executar_receiving_EAD(self):
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
            automation.selecionar_wm_mobile()
            self.df = automation.auto_wm_mobile(self.df)
            self.exibir_dados_planilha()
            messagebox.showinfo("Sucesso", "Execução concluída com sucesso!")
            automation.close_driver()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar função: {e}")
            if 'automation' in locals() and automation.driver:
                automation.close_driver()

if __name__ == "__main__":
    root = ttk.Window(themename="litera")
    root._style = ttk.Style()
    app = TelaPrincipal(root)
    root.mainloop()
