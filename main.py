from tkinter import *
import tkinter as tk
from PIL import Image, ImageTk
import pymysql
from pymysql import MySQLError
import pyodbc
import pandas as pd
from tkinter import filedialog
from os.path import exists

# Cores
cor1 = "#302f2f"  # cinza
cor2 = "#787775"
EXCEL_FILE_PATH = "dados_mosquitos.xlsx"
genero_to_subfamilia_tribo = {
    "Haemagogus": {"subfamilia": "Culicinae", "tribo": "Aedini"},
    "Psorophora": {"subfamilia": "Culicinae", "tribo": "Sabethini"},
    "Aedes": {"subfamilia": "Culicinae", "tribo": "Aedini"},
    "Trichoprosopon": {"subfamilia": "Culicinae", "tribo": "Sabethini"},
    "Sabethes": {"subfamilia": "Culicinae", "tribo": "Sabethini"},
    "Limatus": {"subfamilia": "Culicinae", "tribo": "Sabethini"},
    "Culex": {"subfamilia": "Culicinae", "tribo": "Culicini"}
}

def create_first_widgets(db_host, db_name, db_table):
    global data_obj, date_entry, collector_entry
    data_obj.db_host = db_host
    data_obj.db_name = db_name
    data_obj.db_table = db_table

    frame1.pack_forget()  # Ocultar a tela atual
    frame2 = Frame(root, bg=cor1)  # Crie um novo frame para a próxima tela
    frame2.pack(fill="both", expand=True)
    display_logo1(frame2)

    # Widgets para coletar informações do tubo
    park_label = Label(frame2, text="Parque:", font=("Times 15"), bg=cor1, fg="white")
    park_label.pack(pady=10)
    data_obj.park.set("Mangabeiras")  # Padrão
    park_options = ["Mangabeiras", "Lagoa do nado", "Estação ecologica", "Ursulina", "Jacques cousteau", "Horto"]
    park_option_menu = OptionMenu(frame2, data_obj.park, *park_options)
    park_option_menu.config(font=("Times 15"), bg=cor1, fg="white")
    park_option_menu.pack(pady=10)

    method_label = Label(frame2, text="Método de coleta:", font=("Times 15"), bg=cor1, fg="white")
    method_label.pack(pady=10)
    data_obj.collection_method.set("BG UP")  # Padrão
    method_options = ["BG UP", "BG DOWN", "PUÇA UP", "PUÇA DOWN", "ASPIRADOR"]
    method_option_menu = OptionMenu(frame2, data_obj.collection_method, *method_options)
    method_option_menu.config(font=("Times 15"), bg=cor1, fg="white")
    method_option_menu.pack(pady=10)

    point_label = Label(frame2, text="Ponto:", font=("Times 15"), bg=cor1, fg="white")
    point_label.pack(pady=10)
    data_obj.point.set("1")  # Padrão
    point_options = ["1", "2", "3"]
    point_option_menu = OptionMenu(frame2, data_obj.point, *point_options)
    point_option_menu.config(font=("Times 15"), bg=cor1, fg="white")
    point_option_menu.pack(pady=10)

    time_label = Label(frame2, text="Tempo:", font=("Times 15"), bg=cor1, fg="white")
    time_label.pack(pady=10)
    data_obj.time.set("Manhã")  # Padrão
    time_options = ["Manhã", "Tarde"]
    time_option_menu = OptionMenu(frame2, data_obj.time, *time_options)
    time_option_menu.config(font=("Times 15"), bg=cor1, fg="white")
    time_option_menu.pack(pady=10)

    date_label = Label(frame2, text="Data:", font=("Times 15"), bg=cor1, fg="white")
    date_label.pack(pady=10)
    date_entry = Entry(frame2, width=20, font=("Times 15"), bg=cor2, fg="white")
    date_entry.pack(pady=10)

    collector_label = Label(frame2, text="Coletor:", font=("Times 15"), bg=cor1, fg="white")
    collector_label.pack(pady=10)
    collector_entry = Entry(frame2, width=20, font=("Times 15"), bg=cor2, fg="white")
    collector_entry.pack(pady=10)

    next_button = Button(frame2, text="Próximo", font=("Times 15"), bg=cor2, fg="white",
                         command=lambda: create_second_widgets(frame2))
    next_button.pack(pady=10)

def create_second_widgets(parent_frame):
    parent_frame.pack_forget()  # Ocultar a tela atual
    frame3 = Frame(root, bg=cor1)
    frame3.pack(fill="both", expand=True)
    display_logo1(frame3)

    id_label = Label(frame3, text="Número do mosquito:", font=("Times 30"), bg=cor1, fg="white")
    id_label.pack(pady=10)
    id_var = StringVar()
    id_entry = Entry(frame3, width=20, font=("Times 15"), bg=cor1, fg="white", textvariable=id_var)
    id_entry.pack(pady=10)

    submit_button = Button(frame3, text="Enviar", font=("Times 15"), bg=cor2, fg="white",
                           command=lambda: submit_data(id_var.get(), frame3))
    submit_button.pack(pady=10)
class Data:
    def __init__(self):
        self.park = StringVar()
        self.collection_method = StringVar()
        self.point = StringVar()
        self.time = StringVar()
        self.date = StringVar()
        self.collector = StringVar()
        self.mosquito_id = StringVar()
        self.subfamilia = StringVar()
        self.tribo = StringVar()
        self.genero = StringVar()
        self.especie = StringVar()
        self.db_host = None
        self.db_name = None
        self.db_table = None

    def get_data(self):
        return {
            "id_mosquito": self.mosquito_id.get(),
            "parque": self.park.get(),
            "metodo_coleta": self.collection_method.get(),
            "ponto": self.point.get(),
            "tempo": self.time.get(),
            "data": self.date.get(),
            "coletor": self.collector.get(),
            "subfamilia": self.subfamilia.get(),
            "tribo": self.tribo.get(),
            "genero": self.genero.get(),
            "especie": self.especie.get(),
        }

    def save_to_database(self):
        data = self.get_data()
        driver = '{ODBC Driver 17 for SQL Server}'
        trusted_connection = 'yes'
        encrypt = 'no'
        conn_str = f"DRIVER={driver};SERVER={self.db_host};DATABASE={self.db_name};Trusted_Connection={trusted_connection};Encrypt={encrypt};"
        try:
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            cursor.execute(f"""
                INSERT INTO {self.db_table} (id_mosquito, parque, metodo, ponto, tempo, dia, coletor, subfamilia, tribo, genero, especie)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
            data['id_mosquito'], data['parque'], data['metodo_coleta'], data['ponto'], data['tempo'], data['data'],
            data['coletor'], data['subfamilia'], data['tribo'], data['genero'], data['especie']))
            conn.commit()
            cursor.close()
            conn.close()
            print("Dados salvos com sucesso no banco de dados.")
        except Exception as e:
            print("Erro ao salvar os dados:", e)

    def save_to_excel(self):
        data = self.get_data()
        if exists(EXCEL_FILE_PATH):
            # Load existing Excel file
            df = pd.read_excel(EXCEL_FILE_PATH)
            # Append new data
            df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
        else:
            # Create a new DataFrame and save it to Excel
            df = pd.DataFrame([data], columns=COLUMNS)

        # Save the DataFrame to Excel
        df.to_excel(EXCEL_FILE_PATH, index=False)
        print("Dados salvos com sucesso no arquivo Excel.")


date_entry = None
collector_entry = None
root = Tk()
root.title("TaxKey_alfa_01")
root.geometry("1280x720")
root.config(bg=cor1)
root.iconphoto(False, PhotoImage(file="LogoFinal.png"))
data_obj = Data()
def set_genero_direct(genero, parent):
    global date_entry, collector_entry

    # Coletar valores dos widgets antes de destruí-los
    data_obj.date.set(date_entry.get())
    data_obj.collector.set(collector_entry.get())

    data_obj.genero.set(genero)
    print(f"Gênero definido diretamente: {genero}")

    # Definir subfamília e tribo com base no gênero
    if genero in genero_to_subfamilia_tribo:
        data_obj.subfamilia.set(genero_to_subfamilia_tribo[genero]["subfamilia"])
        data_obj.tribo.set(genero_to_subfamilia_tribo[genero]["tribo"])
        print(f"Subfamília definida: {data_obj.subfamilia.get()}")
        print(f"Tribo definida: {data_obj.tribo.get()}")

    # Verifique se o gênero possui espécies associadas
    generos_com_especies = {
        "Haemagogus": ['Haemagogus leucophoebus','Haemagogus leucocelaenus','Haemagogus celeste','Haemagogus tropicalis','Haemagogus janthinomys','Haemagogus albomaculatus','Haemagogus baresi','Haemagogus spegazzinii'],
        "Psorophora": ['Psorophora saeva', 'Psorophora cilipes', 'Psorophora holmbergii', 'Psorophora ciliata', 'Psorophora pallescens', 'Psorophora lanei', 'Psorophora cyanescens', 'Psorophora johnstonii', 'Psorophora varipes', 'Psorophora discrucians', 'Psorophora johnstoni', 'Psorophora totonaci', 'Psorophora circumflava'],
        "Aedes": ['Aedes aegypti', 'Aedes albopictus', 'Aedes scapularis', 'Aedes gabriel'],
        "Trichoprosopon": [],
        "Sabethes": ['Sabethes petrocchiae','Sabethes albiprivus','Sabethes chloropterus','Sabethes conditus','Sabethes tridentatus','Sabethes glaucodaemon','Sabethes purpureus','Sabethes batesi','Sabethes cyaneus','Sabethes lanei','Sabethes spixi','Sabethes quasicyaneus'],
        "Limatus": [],
        "Culex": [],
    }

    if genero in generos_com_especies and len(generos_com_especies[genero]) > 0:
        # Se o gênero tem espécies, mostrar opções de espécies
        show_especie_options(root)
    else:
        # Se o gênero não tem espécies, vá direto para a tela de confirmação
        create_confirmation_widgets()

    # Destruir o parent após criar os widgets necessários
    parent.destroy()

def set_especie_direct(especie, parent_frame):
    data_obj.especie.set(especie)
    print(f"Espécie definida diretamente: {especie}")

    # Definir subfamília e tribo com base no gênero
    genero = data_obj.genero.get()
    if genero in genero_to_subfamilia_tribo:
        data_obj.subfamilia.set(genero_to_subfamilia_tribo[genero]["subfamilia"])
        data_obj.tribo.set(genero_to_subfamilia_tribo[genero]["tribo"])
        print(f"Subfamília definida: {data_obj.subfamilia.get()}")
        print(f"Tribo definida: {data_obj.tribo.get()}")

    # Salvar os dados no banco de dados ou no Excel
    if save_method.get() == "Database":
        data_obj.save_to_database()
    else:
        data_obj.save_to_excel()

    # Criar a tela de confirmação
    create_confirmation_widgets()

    # Ocultar o parent_frame após criar a tela de confirmação
    parent_frame.pack_forget()

def show_genero_options(parent):
    # Remover widgets antigos do parent frame
    for widget in parent.winfo_children():
        widget.pack_forget()

    # Adicionar a lista de gêneros na mesma tela
    generos = ["Haemagogus","Psorophora","Aedes","Trichoprosopon","Sabethes","Limatus","Culex"]
    genero_var = StringVar(parent)

    genero_label = Label(parent, text="Selecione o gênero:", font=("Times 15"), bg=cor1, fg="white")
    genero_label.pack(pady=10)

    genero_menu = OptionMenu(parent, genero_var, *generos)
    genero_menu.pack(pady=10)

    select_button = Button(parent, text="Selecionar", font=("Times 15"), bg=cor2, fg="white",
                           command=lambda: set_genero_direct(genero_var.get(), parent))
    select_button.pack(pady=10)

def show_especie_options(parent):
    # Certificar-se de que um gênero foi selecionado antes de mostrar as espécies
    if not data_obj.genero.get():
        tk.messagebox.showwarning("Aviso", "Por favor, selecione um gênero primeiro.")
        return

    # Remover widgets antigos do parent frame
    for widget in parent.winfo_children():
        widget.pack_forget()

    # Cria um novo frame para as opções de espécie
    frame_especie = Frame(parent, bg=cor1)
    frame_especie.pack(fill="both", expand=True)

    # Adicionar a lista de espécies na mesma tela
    generos_com_especies = {
        "Haemagogus": ['Haemagogus leucophoebus','Haemagogus leucocelaenus','Haemagogus celeste','Haemagogus tropicalis','Haemagogus janthinomys','Haemagogus albomaculatus','Haemagogus baresi','Haemagogus spegazzinii'],
        "Psorophora": ['Psorophora saeva', 'Psorophora cilipes', 'Psorophora holmbergii', 'Psorophora ciliata', 'Psorophora pallescens', 'Psorophora lanei', 'Psorophora cyanescens', 'Psorophora johnstonii', 'Psorophora varipes', 'Psorophora discrucians', 'Psorophora johnstoni', 'Psorophora totonaci', 'Psorophora circumflava'],
        "Aedes": ['Aedes aegypti', 'Aedes albopictus', 'Aedes scapularis', 'Aedes gabriel'],
        "Trichoprosopon": [],
        "Sabethes": ['Sabethes petrocchiae','Sabethes albiprivus','Sabethes chloropterus','Sabethes conditus','Sabethes tridentatus','Sabethes glaucodaemon','Sabethes purpureus','Sabethes batesi','Sabethes cyaneus','Sabethes lanei','Sabethes spixi','Sabethes quasicyaneus'],
        "Limatus": [],
        "Culex": [],
    }

    especies = generos_com_especies.get(data_obj.genero.get(), [])
    if not especies:
        tk.messagebox.showwarning("Aviso", "Este gênero não possui espécies.")
        return

    especie_var = StringVar(frame_especie)

    especie_label = Label(frame_especie, text="Selecione a espécie:", font=("Times 15"), bg=cor1, fg="white")
    especie_label.pack(pady=10)

    especie_menu = OptionMenu(frame_especie, especie_var, *especies)
    especie_menu.pack(pady=10)

    select_button = Button(frame_especie, text="Selecionar", font=("Times 15"), bg=cor2, fg="white",
                           command=lambda: set_especie_direct(especie_var.get(), frame_especie))
    select_button.pack(pady=10)

    chave_button = Button(frame_especie, text="Ir para a chave de definição", font=("Times 15"), bg=cor2, fg="white",
                          command=lambda: open_chave_definicao(frame_especie))
    chave_button.pack(pady=10)

def create_confirmation_widgets():
    frame_confirm = Frame(root, bg=cor1)
    frame_confirm.pack(fill="both", expand=True)
    display_logo(frame_confirm)

    confirmation_label = Label(frame_confirm, text=f"O mosquito {data_obj.mosquito_id.get()} pertence à: \nSubfamília = {data_obj.subfamilia.get()}\nTribo = {data_obj.tribo.get()}\nGênero = {data_obj.genero.get()}\nEspécie = {data_obj.especie.get()}\nDeseja identificar outro mosquito?", font=("Times 35"), bg=cor1, fg="white", wraplength=600)
    confirmation_label.pack(pady=20)

    button_yes = Button(frame_confirm, text="Sim", font=("Times 30"), bg=cor2, fg="white", command=lambda: identify_another_mosquito_subfamilia(frame_confirm))
    button_yes.pack(side="left", padx=20, pady=10)

    button_no = Button(frame_confirm, text="Não", font=("Times 30"), bg=cor2, fg="white", command=lambda: identify_new_characteristics_subfamilia(frame_confirm))
    button_no.pack(side="right", padx=20, pady=10)

def display_logo(parent):
    # Frame para os botões
    frame_buttons = Frame(parent, bg=cor1)
    frame_buttons.pack(side="bottom", fill="x", pady=20, padx=20, anchor="e")

    # Adicionar os botões "Já sei o gênero" e "Já sei a espécie"
    btn_genero = Button(frame_buttons, text="Já sei o gênero", font=("Times 15"), bg=cor2, fg="white",
                        command=lambda: show_genero_options(parent))
    btn_genero.pack(side="left", padx=16)

    # Exibir o logo abaixo dos botões
    logo_image = Image.open("LogoFinal.png")
    logo_image = logo_image.resize((500, 500))  # Ajuste o tamanho conforme necessário
    logo_photo = ImageTk.PhotoImage(logo_image)
    logo_label = tk.Label(parent, image=logo_photo, bg=cor1)
    logo_label.image = logo_photo
    logo_label.pack(side="left")

def display_logo1(parent):
    logo_image = Image.open("LogoFinal.png")
    logo_image = logo_image.resize((500, 500))  # Ajuste o tamanho conforme necessário
    logo_photo = ImageTk.PhotoImage(logo_image)
    logo_label = tk.Label(parent, image=logo_photo, bg=cor1)
    logo_label.image = logo_photo
    logo_label.pack(side="left")

def test_and_proceed(db_host, db_name, db_table):
    if save_method.get() == "Database":
        try:
            driver = '{ODBC Driver 17 for SQL Server}'
            trusted_connection = 'yes'
            encrypt = 'no'
            conn_str = f"DRIVER={driver};SERVER={db_host};DATABASE={db_name};Trusted_Connection={trusted_connection};Encrypt={encrypt};"
            conn = pyodbc.connect(conn_str)
            conn.close()
            # Se a conexão for bem-sucedida, continue para a próxima tela
            create_first_widgets(db_host, db_name, db_table)
            print("Conexão bem-sucedida")
        except Exception as e:
            print("Erro ao conectar ao banco de dados:", e)
            # Mostre uma mensagem de erro na interface do usuário
            error_label = Label(frame1, text="Erro ao conectar ao banco de dados", font=("Times 15"), bg=cor1, fg="red")
            error_label.pack(pady=10)

        # Salvar informações de conexão do banco de dados
        if data_obj is not None:
            data_obj.db_host = db_host
            data_obj.db_name = db_name
            data_obj.db_table = db_table
    elif save_method.get() == "Excel":
        data_obj.save_to_excel()
        create_first_widgets(db_host, db_name, db_table)

def select_genero():
    # Criar uma nova janela para selecionar o gênero
    genero_window = Toplevel(root)
    genero_window.title("Selecionar Gênero")
    genero_window.geometry("400x400")

    generos = ["Gênero 1", "Gênero 2", "Gênero 3"]  # Substituir com a lista real de gêneros
    genero_var = StringVar(genero_window)

    genero_label = Label(genero_window, text="Selecione o gênero:", font=("Times 15"))
    genero_label.pack(pady=10)

    genero_menu = OptionMenu(genero_window, genero_var, *generos)
    genero_menu.pack(pady=10)

    select_button = Button(genero_window, text="Selecionar", font=("Times 15"),
                           command=lambda: set_genero_direct(genero_var.get(), genero_window))
    select_button.pack(pady=10)

def select_especie():
    # Certificar-se de que um gênero foi selecionado antes de mostrar as espécies
    if not data_obj.genero.get():
        tk.messagebox.showwarning("Aviso", "Por favor, selecione um gênero primeiro.")
        return

    # Criar uma nova janela para selecionar a espécie
    especie_window = Toplevel(root)
    especie_window.title("Selecionar Espécie")
    especie_window.geometry("400x400")

    especies = ["Haemagogus","Psorophora","Aedes","Trichoprosopon","Sabethes","Limatus","Culex"]
    especie_var = StringVar(especie_window)

    especie_label = Label(especie_window, text="Selecione a espécie:", font=("Times 15"))
    especie_label.pack(pady=10)

    especie_menu = OptionMenu(especie_window, especie_var, *especies)
    especie_menu.pack(pady=10)

    select_button = Button(especie_window, text="Selecionar", font=("Times 15"),
                           command=lambda: set_especie_direct(especie_var.get(), especie_window))
    select_button.pack(pady=10)

def create_save_method_widgets():
    global save_method
    save_method_label = Label(frame1, text="Salvar dados em:", font="Times 15", bg=cor1, fg="white")
    save_method_label.pack(pady=10)
    database_radio = Radiobutton(frame1, text="Banco de Dados SQL", variable=save_method, value="Database",
                                 font="Times 15", bg=cor1, fg="white", selectcolor=cor2)
    database_radio.pack(pady=10)
    excel_radio = Radiobutton(frame1, text="Arquivo Excel", variable=save_method, value="Excel", font="Times 15",
                              bg=cor1, fg="white", selectcolor=cor2)
    excel_radio.pack(pady=10)

    db_frame = Frame(frame1, bg=cor1)
    db_frame.pack(pady=10)

    db_host_label = Label(db_frame, text="Servidor:", font=("Times 15"), bg=cor1, fg="white")
    db_host_label.grid(row=0, column=0, padx=5, pady=5)
    db_host_entry = Entry(db_frame, width=20, font=("Times 15"), bg=cor2, fg="white")
    db_host_entry.grid(row=0, column=1, padx=5, pady=5)

    db_table_label = Label(db_frame, text="Nome da Tabela:", font=("Times 15"), bg=cor1, fg="white")
    db_table_label.grid(row=2, column=0, padx=5, pady=5)
    db_table_entry = Entry(db_frame, width=20, font=("Times 15"), bg=cor2, fg="white")
    db_table_entry.grid(row=2, column=1, padx=5, pady=5)

    db_name_label = Label(db_frame, text="Nome do Banco:", font=("Times 15"), bg=cor1, fg="white")
    db_name_label.grid(row=1, column=0, padx=5, pady=5)
    db_name_entry = Entry(db_frame, width=20, font=("Times 15"), bg=cor2, fg="white")
    db_name_entry.grid(row=1, column=1, padx=5, pady=5)

    next_button = Button(frame1, text="Próximo", font=("Times 15"), bg=cor2, fg="white",
                         command=lambda: test_and_proceed(db_host_entry.get(), db_name_entry.get(), db_table_entry.get()))
    next_button.pack(pady=10)

def open_chave_definicao(parent):
    # Primeiro, vamos remover todos os widgets do parent frame
    for widget in parent.winfo_children():
        widget.pack_forget()

    # Agora, vamos verificar qual gênero foi selecionado e redirecionar para a chave apropriada
    genero = data_obj.genero.get()
    ["Haemagogus", "Psorophora", "Aedes", "Trichoprosopon", "Sabethes", "Limatus", "Culex"]
    if genero == "Haemagogus":
        create_chave_E_widgets()
    elif genero == "Psorophora":
        create_chave_F_widgets()
    elif genero == "Aedes":
        create_chave_D_widgets()
    elif genero == "Sabethes":
        create_chave_H_widgets()
    else:
        tk.messagebox.showwarning("Aviso", "Gênero não encontrado ou não possui chave de definição.")
        create_confirmation_widgets()

# Frame para conter os widgets
frame1 = Frame(root, bg=cor1)
frame1.pack(fill="both", expand=True)
display_logo(frame1)
save_method = StringVar(value="Database")

def submit_data(mosquito_id, parent_frame):
    data_obj.mosquito_id.set(mosquito_id)
    parent_frame.pack_forget()
    create_identification_widgets()

def create_identification_widgets():
    frame4 = Frame(root, bg=cor1)
    frame4.pack(fill="both", expand=True)
    display_logo(frame4)

    # Pergunta 1 da Chave A
    A1_label = Label(frame4, text=" Subfamilias - Pergunta 1", font=("Times 20"), bg=cor1, fg="white")
    A1_label.pack(pady=10)
    A1_text = Text(frame4, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    A1_text.insert(INSERT,
                   "A - Tergitos abdominais II-VI desprovidos de escamas na maior parte de sua superfície; palpos geralmente tão longos como a probóscide \nB - Tergitos abdominais II-VI sempre dotados de numerosas escamas cobrindo a sua superfície; palpos de comprimento equivalente a não mais do que três quartos da probóscide, mas geralmente mais curtos do que este apêndice")
    A1_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame4, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_subfamilia("Anophelinae", frame4))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame4, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_A2(frame4))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def set_subfamilia(subfamilia, frame):
    data_obj.subfamilia.set(subfamilia)
    frame.pack_forget()
    print(f"Subfamília definida: {subfamilia}")
    # Salvar os dados no banco de dados ou no Excel
    if save_method.get() == "Database":
        data_obj.save_to_database()
    elif save_method.get() == "Excel":
        data_obj.save_to_excel()
    # Cria a tela de confirmação
    create_confirmation_widgets()

def identify_another_mosquito_subfamilia(frame_confirm):
    frame_confirm.pack_forget()
    data_obj.subfamilia.set("")  # Reseta a subfamília
    data_obj.tribo.set("")  # Reseta a tribo
    data_obj.genero.set("")  # Reseta o gênero
    data_obj.especie.set("")  # Reseta a espécie
    create_second_widgets(frame_confirm)  # Modifique aqui

def identify_new_characteristics_subfamilia(frame_confirm):
    frame_confirm.pack_forget()
    create_first_widgets(data_obj.db_host, data_obj.db_name, data_obj.db_table)

def show_question_A2(frame4):
    frame4.pack_forget()
    frame5 = Frame(root, bg=cor1)
    frame5.pack(fill="both", expand=True)
    display_logo(frame5)

    # Pergunta 2 da Chave A
    A2_label = Label(frame5, text="Subfamilias - Pergunta 2", font=("Times 20"), bg=cor1, fg="white")
    A2_label.pack(pady=10)
    A2_text = Text(frame5, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    A2_text.insert(INSERT,
                   "A - Escutelo arredondado\nB - Escutelo trilobado")
    A2_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame5, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_subfamilia("Toxorhynchitini", frame5))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame5, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_subfamilia_and_continue("Culicinae", frame5))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def set_subfamilia_and_continue(subfamilia, frame5):
    data_obj.subfamilia.set(subfamilia)
    frame5.pack_forget()
    print(f"Subfamília definida: {subfamilia}")
    # Aqui você chama a função para a próxima chave (Chave B)
    create_chave_b_widgets()


def create_chave_b_widgets():
    frame6 = Frame(root, bg=cor1)
    frame6.pack(fill="both", expand=True)
    display_logo(frame6)

    # Pergunta 1 da Chave B
    B1_label = Label(frame6, text="Tribos Culicidae - Pergunta 1", font=("Times 20"), bg=cor1, fg="white")
    B1_label.pack(pady=10)
    B1_text = Text(frame6, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    B1_text.insert(INSERT,
                   "A - Probóscide longa e acentuadamente curva, inferiormente recurva em direção posterior -> Tribo = Toxorhynchitini  | PARE A CHAVE\nB - Probóscide longa ou curta, mas sempre reta e quando curva somente o é na extremidade distal, mesmo assim de maneira moderada -> Vá para a pergunta 2")
    B1_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame6, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo("Toxorhynchitini", frame6))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame6, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_B2(frame6))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def set_tribo(tribo, frame):
    data_obj.tribo.set(tribo)
    frame.pack_forget()
    print(f"Tribo definida: {tribo}")
    # Salvar os dados no banco de dados ou no Excel
    if save_method.get() == "Database":
        data_obj.save_to_database()
    elif save_method.get() == "Excel":
        data_obj.save_to_excel()
    # Cria a tela de confirmação
    create_confirmation_widgets()


def show_question_B2(frame):
    frame.pack_forget()
    frame7 = Frame(root, bg=cor1)
    frame7.pack(fill="both", expand=True)
    display_logo(frame7)

    # Pergunta 2 da Chave B
    B2_label = Label(frame7, text="Tribos Culicidae - Pergunta 2", font=("Times 20"), bg=cor1, fg="white")
    B2_label.pack(pady=10)
    B2_text = Text(frame7, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    B2_text.insert(INSERT,
                   "A - Asa com A1 terminando medialmente em relação à bifurcação de CuA > Tribo = Uranotaeniini  | PARE A CHAVE\nB - Asa com A1 terminando lateralmente em relação à bifurcação de CuA -> Vá para a pergunta 3")
    B2_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame7, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo("Uranotaeniini", frame7))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame7, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_B3(frame7))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_B3(frame):
    frame.pack_forget()
    frame8 = Frame(root, bg=cor1)
    frame8.pack(fill="both", expand=True)
    display_logo(frame8)

    # Pergunta 3 da Chave B
    B3_label = Label(frame8, text="Tribos Culicidae - Pergunta 3", font=("Times 20"), bg=cor1, fg="white")
    B3_label.pack(pady=10)
    B3_text = Text(frame8, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    B3_text.insert(INSERT,
                   "A - Mesoposnoto (Mpn) dotado de cerdas, podendo apresentar escamas; base da coxa posterior (Cx-III) situada ao mesmo nível ou mesmo superiormente à margem dorsal do mesomeron (Msm) > Tribo = Sabethini | IR PARA A CHAVE G\nB - Mesoposnoto (Mpn) glabro; base da coxa posterior (Cx-III) nitidamente situada em nível inferior à margem dorsal do mesomeron (Msm) -> Vá para a pergunta 4")
    B3_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame8, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo_and_continue("Sabethini", frame8, create_chave_g_widgets))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame8, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_B4(frame8))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def set_tribo_and_continue(tribo, frame, next_widget_function):
    data_obj.tribo.set(tribo)
    frame.pack_forget()
    print(f"Tribo definida: {tribo}")
    # Continue to the next key widget
    next_widget_function()


def show_question_B4(frame):
    frame.pack_forget()
    frame9 = Frame(root, bg=cor1)
    frame9.pack(fill="both", expand=True)
    display_logo(frame9)

    # Pergunta 4 da Chave B
    B4_label = Label(frame9, text="Tribos Culicidae - Pergunta 4", font=("Times 20"), bg=cor1, fg="white")
    B4_label.pack(pady=10)
    B4_text = Text(frame9, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    B4_text.insert(INSERT,
                   "A - Bicudo (Esc) revestido de escamas metálicas e sem cerdas pré-escutelares (SPr) > Tribo = Aedini | IR PARA A CHAVE C\nB - Escudo (Esc) com revestimento diferente e presença de cerdas pré-escutelares (SPr) -> Vá para a pergunta 5")
    B4_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame9, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo_and_continue("Aedini", frame9, create_chave_c_widgets))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame9, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_B5(frame9))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_B5(frame):
    frame.pack_forget()
    frame10 = Frame(root, bg=cor1)
    frame10.pack(fill="both", expand=True)
    display_logo(frame10)

    # Pergunta 5 da Chave B
    B5_label = Label(frame10, text="Tribos Culicidae - Pergunta 5", font=("Times 20"), bg=cor1, fg="white")
    B5_label.pack(pady=10)
    B5_text = Text(frame10, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    B5_text.insert(INSERT,
                   "A - Cerdas pós-espiraculares (SAP) presentes -> Vá para a pergunta 6\nB - Cerdas pós-espiraculares (SAP) ausentes -> Vá para a pergunta 7")
    B5_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame10, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_B6(frame10))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame10, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_B7(frame10))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_B6(frame):
    frame.pack_forget()
    frame11 = Frame(root, bg=cor1)
    frame11.pack(fill="both", expand=True)
    display_logo(frame11)

    # Pergunta 6 da Chave B
    B6_label = Label(frame11, text="Tribos Culicidae - Pergunta 6", font=("Times 20"), bg=cor1, fg="white")
    B6_label.pack(pady=10)
    B6_text = Text(frame11, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    B6_text.insert(INSERT,
                   "A - Asa com escamas largas; abdomen da fêmea com extremidade distal truncada > Tribo = Mansoniini | PARE A CHAVE\nB - Asa com escamas estreitas; abdome da fêmea com a extremidade distal afilada e geralmente telescopada > Tribo = Aedini | IR PARA A CHAVE C")
    B6_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame11, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo("Mansoniini", frame11))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame11, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo_and_continue("Aedini", frame11, create_chave_c_widgets))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_B7(frame):
    frame.pack_forget()
    frame12 = Frame(root, bg=cor1)
    frame12.pack(fill="both", expand=True)
    display_logo(frame12)

    # Pergunta 7 da Chave B
    B7_label = Label(frame12, text="Tribos Culicidae - Pergunta 7", font=("Times 20"), bg=cor1, fg="white")
    B7_label.pack(pady=10)
    B7_text = Text(frame12, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    B7_text.insert(INSERT,
                   "A - Asa apresentando, na fase inferior, cerdas implantadas na base da veia subcosta (Se) > Tribo = Culisetini | PARE A CHAVE\nB - Asa desprovida, na face inferior, de cerdas implantadas na base da veia subcosta (Se) -> Vá para a pergunta 8")
    B7_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame12, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo("Culisetini", frame12))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame12, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_B8(frame12))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_B8(frame):
    frame.pack_forget()
    frame12 = Frame(root, bg=cor1)
    frame12.pack(fill="both", expand=True)
    display_logo(frame12)

    # Pergunta 8 da Chave B
    B8_label = Label(frame12, text="Tribos Culicidae - Pergunta 8", font=("Times 20"), bg=cor1, fg="white")
    B8_label.pack(pady=10)
    B8_text = Text(frame12, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    B8_text.insert(INSERT,
                   "A - Antena com flagelômeros (Flm) curtos e grossos; fêmures, médio e posterior, com tufo apical de escamas > Tribo = Aedeomyiini | PARE A CHAVE\nB - Antena com flagelômeros (Fim) alongados e estreitos; fêmures, médio e posterior, sem tufo apical de escamas -> Vá para a pergunta 9")
    B8_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame12, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo("Aedeomyiini", frame12))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame12, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_B9(frame12))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_B9(frame):
    frame.pack_forget()
    frame13 = Frame(root, bg=cor1)
    frame13.pack(fill="both", expand=True)
    display_logo(frame13)

    # Pergunta 9 da Chave B
    B9_label = Label(frame13, text="Tribos Culicidae - Pergunta 9", font=("Times 20"), bg=cor1, fg="white")
    B9_label.pack(pady=10)
    B9_text = Text(frame13, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    B9_text.insert(INSERT,
                   "A - Tarsômeros Ta-4, anterior e médio, curtos e tão longos quanto largos > Tribo = Orthopodomyiini | PARE A CHAVE\nB - Tarsômeros Ta-4, anterior e médio, alongados e mais longos do que largos -> Vá para a pergunta 10")
    B9_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame13, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo("Orthopodomyiini", frame13))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame13, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_B10(frame13))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_B10(frame):
    frame.pack_forget()
    frame14 = Frame(root, bg=cor1)
    frame14.pack(fill="both", expand=True)
    display_logo(frame14)

    # Pergunta 10 da Chave B
    B10_label = Label(frame14, text="Tribos Culicidae - Pergunta 10", font=("Times 20"), bg=cor1, fg="white")
    B10_label.pack(pady=10)
    B10_text = Text(frame14, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    B10_text.insert(INSERT,
                    "A - Ápice do abdome feminino com aspecto afilado ou telescopado; escudo com desenho de escamas douradas > Tribo = Aedini | IR PARA A CHAVE C\nB - Ápice do abdome feminino com aspecto truncado ou rombo; escudo com outro tipo de desenho ou mesmo sem nenhum > Tribo = Culicini | IR PARA A CHAVE I")
    B10_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame14, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo_and_continue("Aedini", frame14, create_chave_c_widgets))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame14, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_tribo_and_continue("Culicini", frame14, create_chave_i_widgets))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def set_genero(genero, frame):
    data_obj.genero.set(genero)
    frame.pack_forget()
    print(f"Genero definido: {genero}")
  # Salvar os dados no banco de dados ou no Excel
    if save_method.get() == "Database":
        data_obj.save_to_database()
    elif save_method.get() == "Excel":
        data_obj.save_to_excel()
    # Cria a tela de confirmação
    create_confirmation_widgets()


def set_genero_and_continue(genero, frame, next_widget_function):
    data_obj.genero.set(genero)
    frame.pack_forget()
    print(f"Genero definido: {genero}")
    next_widget_function()

def create_chave_c_widgets():
    frame15 = Frame(root, bg=cor1)
    frame15.pack(fill="both", expand=True)
    display_logo(frame15)

    # Pergunta 1 da Chave C
    C1_label = Label(frame15, text="Generos Aedini - Pergunta 1", font=("Times 20"), bg=cor1, fg="white")
    C1_label.pack(pady=10)
    C1_text = Text(frame15, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    C1_text.insert(INSERT,
                   "A - Escudo (Esc) revestido de escamas com totalidade metálica > Gênero = Haemagogus\nB - Escudo (Esc) revestido de escamas diferentes > C2")
    C1_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame15, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero_and_continue("Haemagogus", frame15, create_chave_E_widgets))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame15, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_C2(frame15))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def show_question_C2(frame):
    frame.pack_forget()
    frame16 = Frame(root, bg=cor1)
    frame16.pack(fill="both", expand=True)
    display_logo(frame16)

    # Pergunta 2 da Chave C
    C2_label = Label(frame16, text="Genero Aedini - Pergunta 2", font=("Times 20"), bg=cor1, fg="white")
    C2_label.pack(pady=10)
    C2_text = Text(frame16, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    C2_text.insert(INSERT, "A - Cerdas pré-espiraculares (SPe) presentes > Gênero = Psorophora\nB - Cerdas pré-espiraculares (SPe) ausentes > C3")
    C2_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame16, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white", command=lambda:set_genero_and_continue("Psorophora", frame16, create_chave_F_widgets))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame16, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white", command=lambda:show_question_C3(frame16))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def show_question_C3(frame):
    frame.pack_forget()
    frame17 = Frame(root, bg=cor1)
    frame17.pack(fill="both", expand=True)
    display_logo(frame17)

    # Pergunta 3 da Chave C
    C3_label = Label(frame17, text="Genero Aedini - Pergunta 3", font=("Times 20"), bg=cor1, fg="white")
    C3_label.pack(pady=10)
    C3_text = Text(frame17, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    C3_text.insert(INSERT, "A - Pleuras com escamas prateadas dispostas em faixas verticais > Gênero = Haemagogus\nB - Pleuras sem essa ornamentação da vestimenta, embora possa apresentar escamas claras > Gênero = Aedes ")
    C3_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame17, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white", command=lambda: set_genero_and_continue("Haemagogus", frame17, create_chave_E_widgets))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame17, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white", command=lambda: set_genero_and_continue("Aedes", frame17, create_chave_D_widgets))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def set_especie(especie, frame):
    data_obj.especie.set(especie)
    frame.pack_forget()
    print(f"Especie definida: {especie}")
  # Salvar os dados no banco de dados ou no Excel
    if save_method.get() == "Database":
        data_obj.save_to_database()
    elif save_method.get() == "Excel":
        data_obj.save_to_excel()
    # Cria a tela de confirmação
    create_confirmation_widgets()
def create_chave_D_widgets():
    frame18 = Frame(root, bg=cor1)
    frame18.pack(fill="both", expand=True)
    display_logo(frame18)

    # Pergunta 1 da Chave C
    D1_label = Label(frame18, text="Especie Aedes - Pergunta 1", font=("Times 20"), bg=cor1, fg="white")
    D1_label.pack(pady=10)
    D1_text = Text(frame18, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D1_text.insert(INSERT,
                   "A - Escudo (Esc) com escamas claras de tonalidade prateada > 2\nB - Escudo (Esc) com escamas claras de tonalidade branca, esbranquiçada, amarelada ou então destituído de escamas claras > 3")
    D1_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame18, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D2(frame18))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame18, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D3(frame18))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def show_question_D2(frame):
    frame.pack_forget()
    frame19 = Frame(root, bg=cor1)
    frame19.pack(fill="both", expand=True)
    display_logo(frame19)

    # Pergunta 2 da Chave D
    D2_label = Label(frame19, text="Especie Aedes - Pergunta 2", font=("Times 20"), bg=cor1, fg="white")
    D2_label.pack(pady=10)
    D2_text = Text(frame19, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D2_text.insert(INSERT,
                   "A- Clípeo (Clp) com escamas prateadas; escudo (Esc) com desenho prateado em forma de lira\nB- Clípeo (Clp) com escamas escuras; escudo (Esc) com linha prateada percorrendo-o longitudinalmente")
    D2_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame19, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes aegypti",frame19))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame19, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes albopictus",frame19))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def show_question_D3(frame):
    frame.pack_forget()
    frame20 = Frame(root, bg=cor1)
    frame20.pack(fill="both", expand=True)
    display_logo(frame20)

    # Pergunta 3 da Chave D
    D3_label = Label(frame20, text="Especie Aedes - Pergunta 3", font=("Times 20"), bg=cor1, fg="white")
    D3_label.pack(pady=10)
    D3_text = Text(frame20, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D3_text.insert(INSERT, "A- Cerdas pós-espiraculares (SAP), ausentes\nB- Cerdas pós-espiraculares (SAP), presentes")
    D3_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame20, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes purpureipes", frame20))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame20, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D4(frame20))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D4(frame):
    frame.pack_forget()
    frame21 = Frame(root, bg=cor1)
    frame21.pack(fill="both", expand=True)
    display_logo(frame21)

    # Pergunta 4 da Chave D
    D4_label = Label(frame21, text="Especie Aedes - Pergunta 4", font=("Times 20"), bg=cor1, fg="white")
    D4_label.pack(pady=10)
    D4_text = Text(frame21, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D4_text.insert(INSERT,
                   "A- Garras tarsais (G) femininas, anterior e média, ao menos uma delas, denteadas\nB- Garras tarsais (G) femininas, anterior e média, lisas")
    D4_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame21, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D5(frame21))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame21, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D77(frame21))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D5(frame):
    frame.pack_forget()
    frame22 = Frame(root, bg=cor1)
    frame22.pack(fill="both", expand=True)
    display_logo(frame22)

    # Pergunta 5 da Chave D
    D5_label = Label(frame22, text="Especie Aedes - Pergunta 5", font=("Times 20"), bg=cor1, fg="white")
    D5_label.pack(pady=10)
    D5_text = Text(frame22, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D5_text.insert(INSERT,
                   "A- Mesocatepisterno (Mcs) sem escamas no botão pré-alar (BPal)\nB- Mesocatepisterno (Mcs) com escamas no botão pré-alar (BPal)")
    D5_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame22, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes mediovittatus", frame22))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame22, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D6(frame22))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D6(frame):
    frame.pack_forget()
    frame23 = Frame(root, bg=cor1)
    frame23.pack(fill="both", expand=True)
    display_logo(frame23)

    # Pergunta 6 da Chave D
    D6_label = Label(frame23, text="Especie Aedes - Pergunta 6", font=("Times 20"), bg=cor1, fg="white")
    D6_label.pack(pady=10)
    D6_text = Text(frame23, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D6_text.insert(INSERT,
                   "A- Segmento VIII feminino, estreito e retraído no segmento VII, de maneira a não ser visível ao exame externo; cercas longas e estreitas\nB- Segmento VIII feminino, um tanto largo, pouco retrátil, apresentando-se visível ao exame externo, revestido total ou parcialmente de escamas; cercas curtas")
    D6_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame23, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D7(frame23))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame23, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D46(frame23))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D7(frame):
    frame.pack_forget()
    frame24 = Frame(root, bg=cor1)
    frame24.pack(fill="both", expand=True)
    display_logo(frame24)

    # Pergunta 7 da Chave D
    D7_label = Label(frame24, text="Especie Aedes - Pergunta 7", font=("Times 20"), bg=cor1, fg="white")
    D7_label.pack(pady=10)
    D7_text = Text(frame24, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D7_text.insert(INSERT,
                   "A- Asa com escamas claras e escuras, ainda que aquelas se limitem a pequena área na porção basal das veias longitudinais anteriores\nB- Asa com escamas escuras, em sua totalidade")
    D7_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame24, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D8(frame24))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame24, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D14(frame24))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D8(frame):
    frame.pack_forget()
    frame25 = Frame(root, bg=cor1)
    frame25.pack(fill="both", expand=True)
    display_logo(frame25)

    # Pergunta 8 da Chave D
    D8_label = Label(frame25, text="Especie Aedes - Pergunta 8", font=("Times 20"), bg=cor1, fg="white")
    D8_label.pack(pady=10)
    D8_text = Text(frame25, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D8_text.insert(INSERT,
                   "A- Asa com as veias dotadas de escamas claras e escuras, misturadas\nB- Asa com escamas claras limitadas às veias longitudinais anteriores")
    D8_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame25, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D9(frame25))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame25, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D10(frame25))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D9(frame):
    frame.pack_forget()
    frame26 = Frame(root, bg=cor1)
    frame26.pack(fill="both", expand=True)
    display_logo(frame26)

    # Pergunta 9 da Chave D
    D9_label = Label(frame26, text="Especie Aedes - Pergunta 9", font=("Times 20"), bg=cor1, fg="white")
    D9_label.pack(pady=10)
    D9_text = Text(frame26, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D9_text.insert(INSERT,
                   "A- Escudo (Esc) coberto de escamas de tonalidade dourada, clara na porção mediana e bronzeada nas áreas laterais\nB- Escudo (Esc) coberto de escamas de tonalidade bronzeada e dourada, esta formando desenho de linhas longitudinais simétricas")
    D9_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame26, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes sollicitans", frame26))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame26, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes albifasciatus", frame26))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D10(frame):
    frame.pack_forget()
    frame27 = Frame(root, bg=cor1)
    frame27.pack(fill="both", expand=True)
    display_logo(frame27)

    # Pergunta 10 da Chave D
    D10_label = Label(frame27, text="Especie Aedes - Pergunta 10", font=("Times 20"), bg=cor1, fg="white")
    D10_label.pack(pady=10)
    D10_text = Text(frame27, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D10_text.insert(INSERT,
                    "A- Asa com escamas claras dispostas em pequena área na base das veias costa (C) e rádio (R)\nB- Asa com escamas claras, brancas, dispostas ao longo da costa (C) e rádio (R), cobrindo a metade ou mais da extensão dessas veias, contada a partir da base")
    D10_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame27, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D11(frame27))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame27, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D12(frame27))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D11(frame):
    frame.pack_forget()
    frame28 = Frame(root, bg=cor1)
    frame28.pack(fill="both", expand=True)
    display_logo(frame28)

    # Pergunta 11 da Chave D
    D11_label = Label(frame28, text="Especie Aedes - Pergunta 11", font=("Times 20"), bg=cor1, fg="white")
    D11_label.pack(pady=10)
    D11_text = Text(frame28, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D11_text.insert(INSERT,
                    "A- Escudo (Esc) de tegumento marrom escuro e com duas amplas manchas amarelas anterolaterais, cobertas de escamas da mesma tonalidade\nB- Escudo (Esc) de tegumento marrom escuro, mas com três linhas longitudinais de escamas douradas")
    D11_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame28, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes pennai", frame28))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame28, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes laguna", frame28))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D12(frame):
    frame.pack_forget()
    frame29 = Frame(root, bg=cor1)
    frame29.pack(fill="both", expand=True)
    display_logo(frame29)

    # Pergunta 12 da Chave D
    D12_label = Label(frame29, text="Especie Aedes - Pergunta 12", font=("Times 20"), bg=cor1, fg="white")
    D12_label.pack(pady=10)
    D12_text = Text(frame29, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D12_text.insert(INSERT,
                    "A- Abdome com os tergitos cobertos inteiramente de escamas amarelo-douradas\nB- Abdome com os tergitos apicalmente revestidos de escamas escuras e basolateralmente de elementos dourados ou amarelados")
    D12_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame29, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes bimaculatus", frame29))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame29, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D13(frame29))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D13(frame):
    frame.pack_forget()
    frame30 = Frame(root, bg=cor1)
    frame30.pack(fill="both", expand=True)
    display_logo(frame30)

    # Pergunta 13 da Chave D
    D13_label = Label(frame30, text="Especie Aedes - Pergunta 13", font=("Times 20"), bg=cor1, fg="white")
    D13_label.pack(pady=10)
    D13_text = Text(frame30, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D13_text.insert(INSERT,
                    "A- Escudo (Esc) com tegumento apresentando mancha acentuadamente escura, posterolateral, junto à raiz das asas\nB- Escudo (Esc) com tegumento apresentando mancha escura postero-lateral pouco acentuada; faixa pigmentada de tonalidade marrom escura, atingindo o mesanepímero (Mam) e o mesocatepisterno (Mcs)")
    D13_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame30, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes stigmaticus", frame30))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame30, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes fulvus", frame30))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D14(frame):
    frame.pack_forget()
    frame31 = Frame(root, bg=cor1)
    frame31.pack(fill="both", expand=True)
    display_logo(frame31)

    # Pergunta 14 da Chave D
    D14_label = Label(frame31, text="Especie Aedes - Pergunta 14", font=("Times 20"), bg=cor1, fg="white")
    D14_label.pack(pady=10)
    D14_text = Text(frame31, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D14_text.insert(INSERT,
                    "A- Tarsômeros com marcação clara, pelo menos os posteriores, entendendo-se como tal a presença de anéis com essa tonalidade, mas não a existência de possíveis faixas claras longitudinais esmaecidas ou não, na face posterior\nB- Tarsômeros escuros, feita a ressalva constante da dicotomia anterior")
    D14_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame31, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D15(frame31))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame31, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D18(frame31))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D15(frame):
    frame.pack_forget()
    frame32 = Frame(root, bg=cor1)
    frame32.pack(fill="both", expand=True)
    display_logo(frame32)

    # Pergunta 15 da Chave D
    D15_label = Label(frame32, text="Especie Aedes - Pergunta 15", font=("Times 20"), bg=cor1, fg="white")
    D15_label.pack(pady=10)
    D15_text = Text(frame32, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D15_text.insert(INSERT,
                    "A- Probóscide (P) com anel claro mediano, distinto\nB- Probóscide (P) sem anel claro mediano evidente")
    D15_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame32, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D16(frame32))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame32, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D17(frame32))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D16(frame):
    frame.pack_forget()
    frame33 = Frame(root, bg=cor1)
    frame33.pack(fill="both", expand=True)
    display_logo(frame33)

    # Pergunta 16 da Chave D
    D16_label = Label(frame33, text="Especie Aedes - Pergunta 16", font=("Times 20"), bg=cor1, fg="white")
    D16_label.pack(pady=10)
    D16_text = Text(frame33, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D16_text.insert(INSERT,
                    "A- Abdome com tergitos apresentando somente faixas basais de escamas claras\nB- Abdome com tergitos apresentando, além das faixas basais de escamas claras, áreas medianas dessa tonalidade de maneira a formar linha longitudinal contínua na face dorsal dessa parte do corpo")
    D16_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame33, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes taeniorhynchus", frame33))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame33, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes sollicitans", frame33))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D17(frame):
    frame.pack_forget()
    frame34 = Frame(root, bg=cor1)
    frame34.pack(fill="both", expand=True)
    display_logo(frame34)

    # Pergunta 17 da Chave D
    D17_label = Label(frame34, text="Especie Aedes - Pergunta 17", font=("Times 20"), bg=cor1, fg="white")
    D17_label.pack(pady=10)
    D17_text = Text(frame34, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D17_text.insert(INSERT,
                    "A- Tarsômeros com marcação clara evidente nos três tarsos\nB- Tarsômeros com marcação clara mais evidente no tarso posterior, não apresentada nos outros tarsômeros ou apenas em alguns deles")
    D17_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame34, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes milleri", frame34))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame34, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes bejaranoi", frame34))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D18(frame):
    frame.pack_forget()
    frame35 = Frame(root, bg=cor1)
    frame35.pack(fill="both", expand=True)
    display_logo(frame35)

    # Pergunta 18 da Chave D
    D18_label = Label(frame35, text="Especie Aedes - Pergunta 18", font=("Times 20"), bg=cor1, fg="white")
    D18_label.pack(pady=10)
    D18_text = Text(frame35, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D18_text.insert(INSERT, "A- Tegumento torácico amarelo\nB- Tegumento torácico de tonalidade marrom escuro")
    D18_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame35, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D19(frame35))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame35, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D20(frame35))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D19(frame):
    frame.pack_forget()
    frame36 = Frame(root, bg=cor1)
    frame36.pack(fill="both", expand=True)
    display_logo(frame36)

    # Pergunta 19 da Chave D
    D19_label = Label(frame36, text="Especie Aedes - Pergunta 19", font=("Times 20"), bg=cor1, fg="white")
    D19_label.pack(pady=10)
    D19_text = Text(frame36, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D19_text.insert(INSERT,
                    "A- Escudo (Esc) com numerosas escamas amareladas esparsas, sem formarem mancha definida\nB- Escudo (Esc) com duas faixas medianas de escamas escuras, além de mancha posterolateral da mesma tonalidade, junto à raiz alar")
    D19_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame36, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes hortator", frame36))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame36, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes perventor", frame36))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D20(frame):
    frame.pack_forget()
    frame37 = Frame(root, bg=cor1)
    frame37.pack(fill="both", expand=True)
    display_logo(frame37)

    # Pergunta 20 da Chave D
    D20_label = Label(frame37, text="Especie Aedes - Pergunta 20", font=("Times 20"), bg=cor1, fg="white")
    D20_label.pack(pady=10)
    D20_text = Text(frame37, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D20_text.insert(INSERT,
                    "A- Abdome com tergitos destituídos de faixas basais de escamas claras contínuas, somente formando áreas basolaterais dessa tonalidade\nB- Abdome com tergitos apresentando faixas basais de escamas claras contínuas, pelo menos em alguns deles")
    D20_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame37, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D21(frame37))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame37, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D28(frame37))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D21(frame):
    frame.pack_forget()
    frame38 = Frame(root, bg=cor1)
    frame38.pack(fill="both", expand=True)
    display_logo(frame38)

    # Pergunta 21 da Chave D
    D21_label = Label(frame38, text="Especie Aedes - Pergunta 21", font=("Times 20"), bg=cor1, fg="white")
    D21_label.pack(pady=10)
    D21_text = Text(frame38, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D21_text.insert(INSERT,
                    "A- Escudo (Esc) com escamas claras dispostas em faixas longitudinais ou em manchas, formando desenho\nB- Escudo (Esc) totalmente coberto de escamas escuras")
    D21_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame38, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D22(frame38))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame38, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes serratus", frame38))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D22(frame):
    frame.pack_forget()
    frame39 = Frame(root, bg=cor1)
    frame39.pack(fill="both", expand=True)
    display_logo(frame39)

    # Pergunta 22 da Chave D
    D22_label = Label(frame39, text="Especie Aedes - Pergunta 22", font=("Times 20"), bg=cor1, fg="white")
    D22_label.pack(pady=10)
    D22_text = Text(frame39, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D22_text.insert(INSERT,
                    "A- Escudo (Esc) com duas faixas longitudinais paramedianas de escamas claras\nB- Escudo (Esc) com escamas claras medianas, formando faixa ou mancha")
    D22_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame39, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes angustivittatus", frame39))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame39, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D23(frame39))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D23(frame):
    frame.pack_forget()
    frame40 = Frame(root, bg=cor1)
    frame40.pack(fill="both", expand=True)
    display_logo(frame40)

    # Pergunta 23 da Chave D
    D23_label = Label(frame40, text="Especie Aedes - Pergunta 23", font=("Times 20"), bg=cor1, fg="white")
    D23_label.pack(pady=10)
    D23_text = Text(frame40, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D23_text.insert(INSERT,
                    "A- Escudo (Esc) com as escamas claras dispostas formando faixa mediana longitudinal\nB- Escudo (Esc) com as escamas claras dispostas em mancha conspícua ocupando a parte anterior da área")
    D23_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame40, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D24(frame40))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame40, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D27(frame40))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D24(frame):
    frame.pack_forget()
    frame41 = Frame(root, bg=cor1)
    frame41.pack(fill="both", expand=True)
    display_logo(frame41)

    # Pergunta 24 da Chave D
    D24_label = Label(frame41, text="Especie Aedes - Pergunta 24", font=("Times 20"), bg=cor1, fg="white")
    D24_label.pack(pady=10)
    D24_text = Text(frame41, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D24_text.insert(INSERT,
                    "A- Lobo mediano do escutelo (Eto) com escamas douradas ou levemente bronzeadas\nB- Lobo mediano do escutelo (Eto) com escamas brancas")
    D24_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame41, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes serratus", frame41))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame41, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D25(frame41))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D25(frame):
    frame.pack_forget()
    frame42 = Frame(root, bg=cor1)
    frame42.pack(fill="both", expand=True)
    display_logo(frame42)

    # Pergunta 25 da Chave D
    D25_label = Label(frame42, text="Especie Aedes - Pergunta 25", font=("Times 20"), bg=cor1, fg="white")
    D25_label.pack(pady=10)
    D25_text = Text(frame42, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D25_text.insert(INSERT,
                    "A- Espécie com indivíduos de porte médio ou grande, com envergadura alar feminina de 3,0 mm ou mais\nB- Espécie com indivíduos de porte pequeno, com envergadura alar feminina de 2,5 mm ou menos")
    D25_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame42, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes tormentor", frame42))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame42, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D26(frame42))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D26(frame):
    frame.pack_forget()
    frame43 = Frame(root, bg=cor1)
    frame43.pack(fill="both", expand=True)
    display_logo(frame43)

    # Pergunta 26 da Chave D
    D26_label = Label(frame43, text="Especie Aedes - Pergunta 26", font=("Times 20"), bg=cor1, fg="white")
    D26_label.pack(pady=10)
    D26_text = Text(frame43, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D26_text.insert(INSERT,
                    "A- Occipício (Occ) com faixa central de escamas brancas estreitas e limitadas submedianamente por conjunto de outras, largas de tonalidade análoga\nB- Occipício (Occ) com a mesma faixa central, porém limitada analogamente por conjunto de escamas largas e escuras")
    D26_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame43, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes dupreei", frame43))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame43, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes serratus", frame43))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D27(frame):
    frame.pack_forget()
    frame44 = Frame(root, bg=cor1)
    frame44.pack(fill="both", expand=True)
    display_logo(frame44)

    # Pergunta 27 da Chave D
    D27_label = Label(frame44, text="Especie Aedes - Pergunta 27", font=("Times 20"), bg=cor1, fg="white")
    D27_label.pack(pady=10)
    D27_text = Text(frame44, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D27_text.insert(INSERT,
                    "A- Escudo (Esc) com a mancha formada pelas escamas claras, apresentando contornos irregulares e angulosos\nB- Escudo (Esc) com a mancha formada pelas escamas claras, apresentando contornos regulares e arredondados")
    D27_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame44, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes rhyacophylus", frame44))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame44, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes scapularis", frame44))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D28(frame):
    frame.pack_forget()
    frame45 = Frame(root, bg=cor1)
    frame45.pack(fill="both", expand=True)
    display_logo(frame45)

    # Pergunta 28 da Chave D
    D28_label = Label(frame45, text="Especie Aedes - Pergunta 28", font=("Times 20"), bg=cor1, fg="white")
    D28_label.pack(pady=10)
    D28_text = Text(frame45, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D28_text.insert(INSERT,
                    "A- Escudo (Esc) coberto de escamas de tonalidade marrom na parte central e amarelada na lateral da área\nB- Escudo (Esc) coberto de escamas formando faixas ou manchas claras")
    D28_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame45, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes obturbator", frame45))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame45, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D29(frame45))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D29(frame):
    frame.pack_forget()
    frame46 = Frame(root, bg=cor1)
    frame46.pack(fill="both", expand=True)
    display_logo(frame46)

    # Pergunta 29 da Chave D
    D29_label = Label(frame46, text="Especie Aedes - Pergunta 29", font=("Times 20"), bg=cor1, fg="white")
    D29_label.pack(pady=10)
    D29_text = Text(frame46, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D29_text.insert(INSERT,
                    "A- Escudo (Esc) com desenho formado por uma ou mais manchas claras\nB- Escudo (Esc) com desenho formado por uma ou mais faixas claras longitudinais")
    D29_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame46, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D30(frame46))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame46, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D38(frame46))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D30(frame):
    frame.pack_forget()
    frame47 = Frame(root, bg=cor1)
    frame47.pack(fill="both", expand=True)
    display_logo(frame47)

    # Pergunta 30 da Chave D
    D30_label = Label(frame47, text="Especie Aedes - Pergunta 30", font=("Times 20"), bg=cor1, fg="white")
    D30_label.pack(pady=10)
    D30_text = Text(frame47, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D30_text.insert(INSERT,
                    "A- Tíbia posterior apresentando a superfície anterior branca, a qual pode circundar esse segmento; probóscide (P) ventralmente clara\nB- Tíbia posterior inteiramente escura; probóscide (P) inteiramente escura")
    D30_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame47, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D31(frame47))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame47, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D32(frame47))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D31(frame):
    frame.pack_forget()
    frame48 = Frame(root, bg=cor1)
    frame48.pack(fill="both", expand=True)
    display_logo(frame48)

    # Pergunta 31 da Chave D
    D31_label = Label(frame48, text="Especie Aedes - Pergunta 31", font=("Times 20"), bg=cor1, fg="white")
    D31_label.pack(pady=10)
    D31_text = Text(frame48, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D31_text.insert(INSERT,
                    "A- Escudo (E.se) com a parte anterior coberta de escamas predominantemente bronzeadas, tendendo à tonalidade amarelada na parte posterior dessa área\nB- Escudo (Esc) com a parte anterior coberta de escamas brancas, às vezes amareladas, tendendo à tonalidade bronzeada na parte lateral dessa área")
    D31_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame48, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes phaeonotus", frame48))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame48, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes scapularis", frame48))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D32(frame):
    frame.pack_forget()
    frame49 = Frame(root, bg=cor1)
    frame49.pack(fill="both", expand=True)
    display_logo(frame49)

    # Pergunta 32 da Chave D
    D32_label = Label(frame49, text="Especie Aedes - Pergunta 32", font=("Times 20"), bg=cor1, fg="white")
    D32_label.pack(pady=10)
    D32_text = Text(frame49, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D32_text.insert(INSERT,
                    "A- Escudo (Esc) inteiramente coberto de escamas claras, amareladas\nB- Escudo (Esc) com escamas claras e escuras")
    D32_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame49, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes thelcter", frame49))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame49, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D33(frame49))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D33(frame):
    frame.pack_forget()
    frame50 = Frame(root, bg=cor1)
    frame50.pack(fill="both", expand=True)
    display_logo(frame50)

    # Pergunta 33 da Chave D
    D33_label = Label(frame50, text="Especie Aedes - Pergunta 33", font=("Times 20"), bg=cor1, fg="white")
    D33_label.pack(pady=10)
    D33_text = Text(frame50, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D33_text.insert(INSERT,
                    "A- Fêmur posterior escuro, destituído de ápice claro na articulação com a tíbia correspondente\nB- Fêmur posterior com ápice claro na articulação com a tíbia correspondente")
    D33_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame50, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes infirmatus", frame50))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame50, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D34(frame50))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D34(frame):
    frame.pack_forget()
    frame51 = Frame(root, bg=cor1)
    frame51.pack(fill="both", expand=True)
    display_logo(frame51)

    # Pergunta 34 da Chave D
    D34_label = Label(frame51, text="Especie Aedes - Pergunta 34", font=("Times 20"), bg=cor1, fg="white")
    D34_label.pack(pady=10)
    D34_text = Text(frame51, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D34_text.insert(INSERT,
                    "A- Escudo (Esc) com mancha ou pequenas manchas de escamas claras, formando desenho indistinto de escamas claras, douradas, brancas ou mesmo bronzeadas\nB- Escudo (Esc) com mancha, apreciavelmente delineada formada por escamas brancas em sua totalidade")
    D34_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame51, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes tortilis", frame51))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame51, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D35(frame51))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D35(frame):
    frame.pack_forget()
    frame52 = Frame(root, bg=cor1)
    frame52.pack(fill="both", expand=True)
    display_logo(frame52)

    # Pergunta 35 da Chave D
    D35_label = Label(frame52, text="Especie Aedes - Pergunta 35", font=("Times 20"), bg=cor1, fg="white")
    D35_label.pack(pady=10)
    D35_text = Text(frame52, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D35_text.insert(INSERT,
                    "A- Escudo (Esc) com a mancha anterior de escamas brancas dividida longitudinalmente por linha de escamas bronzeadas\nB- Escudo (Esc) sem linha longitudinal de escamas bronzeadas dividindo a mancha clara")
    D35_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame52, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes raymondi", frame52))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame52, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D36(frame52))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D36(frame):
    frame.pack_forget()
    frame53 = Frame(root, bg=cor1)
    frame53.pack(fill="both", expand=True)
    display_logo(frame53)

    # Pergunta 36 da Chave D
    D36_label = Label(frame53, text="Especie Aedes - Pergunta 36", font=("Times 20"), bg=cor1, fg="white")
    D36_label.pack(pady=10)
    D36_text = Text(frame53, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D36_text.insert(INSERT,
                    "A- Cabeça com as escamas erectas numerosas e alongadas, estendendo-se em direção anterior até o vértice(v), e escuras, em sua maioria\nB- Cabeça com escamas eretas de aspecto diferente")
    D36_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame53, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes pectinatus", frame53))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame53, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D37(frame53))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D37(frame):
    frame.pack_forget()
    frame54 = Frame(root, bg=cor1)
    frame54.pack(fill="both", expand=True)
    display_logo(frame54)

    # Pergunta 37 da Chave D
    D37_label = Label(frame54, text="Especie Aedes - Pergunta 37", font=("Times 20"), bg=cor1, fg="white")
    D37_label.pack(pady=10)
    D37_text = Text(frame54, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D37_text.insert(INSERT,
                    "A- Palpos maxilares (PlpM) aparentemente com 3 (três) palpômeros; espécie centro e norte sul-americana\nB- Palpos maxilares (PlpM) aparentemente com 4 (quatro) palpômeros; espécie meridional sul-americana")
    D37_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame54, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes euplocanius", frame54))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame54, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes patersoni", frame54))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D38(frame):
    frame.pack_forget()
    frame55 = Frame(root, bg=cor1)
    frame55.pack(fill="both", expand=True)
    display_logo(frame55)

    # Pergunta 38 da Chave D
    D38_label = Label(frame55, text="Especie Aedes - Pergunta 38", font=("Times 20"), bg=cor1, fg="white")
    D38_label.pack(pady=10)
    D38_text = Text(frame55, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D38_text.insert(INSERT,
                    "A- Escudo (Esc) ornamentado com uma só faixa longitudinal, clara, mediana e larga\nB- Escudo (Esc) ornamentado com mais de uma faixa longitudinal clara")
    D38_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame55, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D39(frame55))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame55, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D41(frame55))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D39(frame):
    frame.pack_forget()
    frame56 = Frame(root, bg=cor1)
    frame56.pack(fill="both", expand=True)
    display_logo(frame56)

    # Pergunta 39 da Chave D
    D39_label = Label(frame56, text="Especie Aedes - Pergunta 39", font=("Times 20"), bg=cor1, fg="white")
    D39_label.pack(pady=10)
    D39_text = Text(frame56, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D39_text.insert(INSERT,
                    "A- Fileira de cerdas dorsocentrais (SD), bem evidente; abdome com os esternitos dotados de linha mediana longitudinal de escamas escuras\nB- Fileira de cerdas dorsocentrais (SD), ausente ou pouco evidente; abdome com os esternitos inteiramente escuros ou com escamas escuras não formando linha longitudinal mediana")
    D39_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame56, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes bogotanus", frame56))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame56, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D40(frame56))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D40(frame):
    frame.pack_forget()
    frame57 = Frame(root, bg=cor1)
    frame57.pack(fill="both", expand=True)
    display_logo(frame57)

    # Pergunta 40 da Chave D
    D40_label = Label(frame57, text="Especie Aedes - Pergunta 40", font=("Times 20"), bg=cor1, fg="white")
    D40_label.pack(pady=10)
    D40_text = Text(frame57, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D40_text.insert(INSERT,
                    "A- Escudo (Esc) com a faixa clara terminando no limite com a área pré-escutelar (Apr)\nB- Escudo (Esc) com a faixa clara contornando lateralmente a área pré-escutelar (Apr) e estendendo-se até o escutelo (Eto), de maneira a este apresentar três conjuntos de escamas claras, cada uma correspondente a cada lobo")
    D40_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame57, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes condolescens", frame57))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame57, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes synchitus", frame57))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D41(frame):
    frame.pack_forget()
    frame58 = Frame(root, bg=cor1)
    frame58.pack(fill="both", expand=True)
    display_logo(frame58)

    # Pergunta 41 da Chave D
    D41_label = Label(frame58, text="Especie Aedes - Pergunta 41", font=("Times 20"), bg=cor1, fg="white")
    D41_label.pack(pady=10)
    D41_text = Text(frame58, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D41_text.insert(INSERT,
                    "A- Escamas escuras do occipício (0cc) formando mancha proeminente\nB- Escamas escuras do occipício (0cc) sem constituírem mancha proeminente")
    D41_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame58, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes incomptus", frame58))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame58, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D42(frame58))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D42(frame):
    frame.pack_forget()
    frame59 = Frame(root, bg=cor1)
    frame59.pack(fill="both", expand=True)
    display_logo(frame59)

    # Pergunta 42 da Chave D
    D42_label = Label(frame59, text="Especie Aedes - Pergunta 42", font=("Times 20"), bg=cor1, fg="white")
    D42_label.pack(pady=10)
    D42_text = Text(frame59, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D42_text.insert(INSERT,
                    "A- Escudo (Esc) com linhas longitudinais de escamas escuras, separando nesse sentido três faixas também longitudinais de escamas claras, estas como componentes de ampla área clara anterior do escudo (Esc)\nB- Escudo (Esc) com duas faixas claras, brancas ou amareladas, na área dorsocentral (AD)")
    D42_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame59, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes crinifer", frame59))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame59, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D43(frame59))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D43(frame):
    frame.pack_forget()
    frame60 = Frame(root, bg=cor1)
    frame60.pack(fill="both", expand=True)
    display_logo(frame60)

    # Pergunta 43 da Chave D
    D43_label = Label(frame60, text="Especie Aedes - Pergunta 43", font=("Times 20"), bg=cor1, fg="white")
    D43_label.pack(pady=10)
    D43_text = Text(frame60, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D43_text.insert(INSERT,
                    "A- Articulações fêmuro-tibiais, principalmente a posterior, com escamas claras\nB- Articulações fêmuro-tibiais, totalmente escuras")
    D43_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame60, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D44(frame60))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame60, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D45(frame60))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D44(frame):
    frame.pack_forget()
    frame61 = Frame(root, bg=cor1)
    frame61.pack(fill="both", expand=True)
    display_logo(frame61)

    # Pergunta 44 da Chave D
    D44_label = Label(frame61, text="Especie Aedes - Pergunta 44", font=("Times 20"), bg=cor1, fg="white")
    D44_label.pack(pady=10)
    D44_text = Text(frame61, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D44_text.insert(INSERT,
                    "A- Linhas de escamas claras dorsocentrais (ED), contínuas até o escutelo (Eto)\nB- Linhas de escamas claras dorsocentrais (ED), interrompidas")
    D44_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame61, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes trivittatus", frame61))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame61, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes atactivittatus", frame61))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D45(frame):
    frame.pack_forget()
    frame62 = Frame(root, bg=cor1)
    frame62.pack(fill="both", expand=True)
    display_logo(frame62)

    # Pergunta 45 da Chave D
    D45_label = Label(frame62, text="Especie Aedes - Pergunta 45", font=("Times 20"), bg=cor1, fg="white")
    D45_label.pack(pady=10)
    D45_text = Text(frame62, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D45_text.insert(INSERT,
                    "A- Linhas de escamas claras dorsocentrais (ED), contínuas e sem se expandir com as escamas supra-alares (ESa)\nB- Linhas de escamas claras dorsocentrais (ED), contínuas mas expandindo-se com as escamas supra-alares (ESa)")
    D45_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame62, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes angustivittatus", frame62))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame62, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes meprai", frame62))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D46(frame):
    frame.pack_forget()
    frame63 = Frame(root, bg=cor1)
    frame63.pack(fill="both", expand=True)
    display_logo(frame63)

    # Pergunta 46 da Chave D
    D46_label = Label(frame63, text="Especie Aedes - Pergunta 46", font=("Times 20"), bg=cor1, fg="white")
    D46_label.pack(pady=10)
    D46_text = Text(frame63, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D46_text.insert(INSERT,
                    "A- Asa com escamas claras e escuras, aquelas limitadas a pequenos ou médios conjuntos localizados na base das veias costa ( C) e rádio (R)\nB- Asa com escamas escuras, em sua totalidade")
    D46_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame63, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D47(frame63))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame63, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D51(frame63))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D47(frame):
    frame.pack_forget()
    frame64 = Frame(root, bg=cor1)
    frame64.pack(fill="both", expand=True)
    display_logo(frame64)

    # Pergunta 47 da Chave D
    D47_label = Label(frame64, text="Especie Aedes - Pergunta 47", font=("Times 20"), bg=cor1, fg="white")
    D47_label.pack(pady=10)
    D47_text = Text(frame64, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D47_text.insert(INSERT,
                    "A- Palpos maxilares (PlpM) femininos com escamas claras\nB- Palpos maxilares (PlpM) femininos, inteiramente escuros")
    D47_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame64, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes mediovittatus", frame64))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame64, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D48(frame64))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D48(frame):
    frame.pack_forget()
    frame65 = Frame(root, bg=cor1)
    frame65.pack(fill="both", expand=True)
    display_logo(frame65)

    # Pergunta 48 da Chave D
    D48_label = Label(frame65, text="Especie Aedes - Pergunta 48", font=("Times 20"), bg=cor1, fg="white")
    D48_label.pack(pady=10)
    D48_text = Text(frame65, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D48_text.insert(INSERT,
                    "A- Tarsos posteriores, com os tarsômeros Ta-3 a 5 apresentando marcação clara basal\nB- Tarsos posteriores, com os tarsômeros Ta-3 a 5 inteiramente escuros")
    D48_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame65, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D49(frame65))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame65, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D50(frame65))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D49(frame):
    frame.pack_forget()
    frame66 = Frame(root, bg=cor1)
    frame66.pack(fill="both", expand=True)
    display_logo(frame66)

    # Pergunta 49 da Chave D
    D49_label = Label(frame66, text="Especie Aedes - Pergunta 49", font=("Times 20"), bg=cor1, fg="white")
    D49_label.pack(pady=10)
    D49_text = Text(frame66, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D49_text.insert(INSERT,
                    "A- Marcação clara basal visível somente nos tarsômeros Ta 3 a 5, sendo o último (Ta-5) a mais das vezes inteiramente claro\nB- Marcação clara basal visível em todos os tarsômeros")
    D49_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame66, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes epactius", frame66))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame66, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("", frame66))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D50(frame):
    frame.pack_forget()
    frame67 = Frame(root, bg=cor1)
    frame67.pack(fill="both", expand=True)
    display_logo(frame67)

    # Pergunta 50 da Chave D
    D50_label = Label(frame67, text="Especie Aedes - Pergunta 50", font=("Times 20"), bg=cor1, fg="white")
    D50_label.pack(pady=10)
    D50_text = Text(frame67, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D50_text.insert(INSERT,
                    "A- Tarso posterior totalmente escuro e destituído de marcação clara na articulação dos tarsômeros Ta 1 e 2\nB- Tarso posterior com área clara na articulação dos tarsômeros Ta 1 e 2")
    D50_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame67, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes upatensis", frame67))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame67, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D51(frame67))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D51(frame):
    frame.pack_forget()
    frame68 = Frame(root, bg=cor1)
    frame68.pack(fill="both", expand=True)
    display_logo(frame68)

    # Pergunta 51 da Chave D
    D51_label = Label(frame68, text="Especie Aedes - Pergunta 51", font=("Times 20"), bg=cor1, fg="white")
    D51_label.pack(pady=10)
    D51_text = Text(frame68, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D51_text.insert(INSERT,
                    "A- Fêmur posterior com conspícuo anel escuro basal e a mácula da área supra-alar (ASa) atingindo a sutura escudo-escutelo (see)\nB- Ambas as condições acima descritas não ocorrendo simultaneamente")
    D51_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame68, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D52(frame68))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame68, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D61(frame68))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D52(frame):
    frame.pack_forget()
    frame69 = Frame(root, bg=cor1)
    frame69.pack(fill="both", expand=True)
    display_logo(frame69)

    # Pergunta 52 da Chave D
    D52_label = Label(frame69, text="Especie Aedes - Pergunta 52", font=("Times 20"), bg=cor1, fg="white")
    D52_label.pack(pady=10)
    D52_text = Text(frame69, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D52_text.insert(INSERT,
                    "A- Tarsômeros médios e posteriores, com marcação basal clara bem discreta ou imperceptível\nB- Tarsômeros médios e posteriores, com marcação basal clara evidente")
    D52_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame69, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes buenaventura", frame69))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame69, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D53(frame69))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D53(frame):
    frame.pack_forget()
    frame70 = Frame(root, bg=cor1)
    frame70.pack(fill="both", expand=True)
    display_logo(frame70)

    # Pergunta 53 da Chave D
    D53_label = Label(frame70, text="Especie Aedes - Pergunta 53", font=("Times 20"), bg=cor1, fg="white")
    D53_label.pack(pady=10)
    D53_text = Text(frame70, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D53_text.insert(INSERT,
                    "A- Tarsômero mediano Ta-5, claro\nB- Tarsômero mediano Ta-5, escuro; o Ta-5 posterior, sempre escuro")
    D53_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame70, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes alboapicus", frame70))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame70, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D54(frame70))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D54(frame):
    frame.pack_forget()
    frame71 = Frame(root, bg=cor1)
    frame71.pack(fill="both", expand=True)
    display_logo(frame71)

    # Pergunta 54 da Chave D
    D54_label = Label(frame71, text="Especie Aedes - Pergunta 54", font=("Times 20"), bg=cor1, fg="white")
    D54_label.pack(pady=10)
    D54_text = Text(frame71, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D54_text.insert(INSERT,
                    "A- As escamas deitadas dispostas ao longo da sutura coronária (se) do vértice (V), todas largas\nB- As escamas deitadas dispostas ao longo da sutura coronária (se) do vértice ( V), todas estreitas e curvas")
    D54_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame71, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D55(frame71))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame71, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D56(frame71))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D55(frame):
    frame.pack_forget()
    frame72 = Frame(root, bg=cor1)
    frame72.pack(fill="both", expand=True)
    display_logo(frame72)

    # Pergunta 55 da Chave D
    D55_label = Label(frame72, text="Especie Aedes - Pergunta 55", font=("Times 20"), bg=cor1, fg="white")
    D55_label.pack(pady=10)
    D55_text = Text(frame72, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D55_text.insert(INSERT,
                    "A- Escudo (Esc) com área clara transversal na parte anterior\nB- Escudo (Esc) sem área clara transversal na parte anterior; poucas ou raras escamas brancas esparsas pela superfície restante do escudo (Esc) e da área pré-escutelar (Apr)")
    D55_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame72, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes thorntoni", frame72))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame72, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes argyrothorax", frame72))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D56(frame):
    frame.pack_forget()
    frame73 = Frame(root, bg=cor1)
    frame73.pack(fill="both", expand=True)
    display_logo(frame73)

    # Pergunta 56 da Chave D
    D56_label = Label(frame73, text="Especie Aedes - Pergunta 56", font=("Times 20"), bg=cor1, fg="white")
    D56_label.pack(pady=10)
    D56_text = Text(frame73, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D56_text.insert(INSERT,
                    "A- Escamas subespiraculares (EASe), ausentes\nB- Escamas subespiraculares (EASe), presentes")
    D56_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame73, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D57(frame73))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame73, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D58(frame73))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D57(frame):
    frame.pack_forget()
    frame74 = Frame(root, bg=cor1)
    frame74.pack(fill="both", expand=True)
    display_logo(frame74)

    # Pergunta 57 da Chave D
    D57_label = Label(frame74, text="Especie Aedes - Pergunta 57", font=("Times 20"), bg=cor1, fg="white")
    D57_label.pack(pady=10)
    D57_text = Text(frame74, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D57_text.insert(INSERT,
                    "A- Vértice (V) com todas as escamas deitadas, estreitas e curvas, a maioria ao longo da sutura coronária (se)\nB- Vértice (V) com área na qual observa-se a presença de escamas deitadas largas, situado adjacente à linha de escamas da sutura coronária (se) que se mostram claras")
    D57_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame74, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes metoecopus", frame74))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame74, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes berlirni", frame74))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D58(frame):
    frame.pack_forget()
    frame75 = Frame(root, bg=cor1)
    frame75.pack(fill="both", expand=True)
    display_logo(frame75)

    # Pergunta 58 da Chave D
    D58_label = Label(frame75, text="Especie Aedes - Pergunta 58", font=("Times 20"), bg=cor1, fg="white")
    D58_label.pack(pady=10)
    D58_text = Text(frame75, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D58_text.insert(INSERT,
                    "A- Vértice (V) com todas as escamas deitadas, estreitas e curvas\nB- Vértice (V) com área na qual observa-se a presença de escamas deitadas largas, situado adjacente à linha de escamas da sutura coronária (sc)")
    D58_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame75, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D59(frame75))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame75, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D60(frame75))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D59(frame):
    frame.pack_forget()
    frame76 = Frame(root, bg=cor1)
    frame76.pack(fill="both", expand=True)
    display_logo(frame76)

    # Pergunta 59 da Chave D
    D59_label = Label(frame76, text="Especie Aedes - Pergunta 59", font=("Times 20"), bg=cor1, fg="white")
    D59_label.pack(pady=10)
    D59_text = Text(frame76, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D59_text.insert(INSERT, "A- Cerdas pré-alares (SPal), escuras\nB- Cerdas pré-alares (SPal), claras")
    D59_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame76, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes braziliensis", frame76))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame76, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes insolitus", frame76))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D60(frame):
    frame.pack_forget()
    frame77 = Frame(root, bg=cor1)
    frame77.pack(fill="both", expand=True)
    display_logo(frame77)

    # Pergunta 60 da Chave D
    D60_label = Label(frame77, text="Especie Aedes - Pergunta 60", font=("Times 20"), bg=cor1, fg="white")
    D60_label.pack(pady=10)
    D60_text = Text(frame77, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D60_text.insert(INSERT,
                    "A- Fêmur posterior com faixa escura basal estreita (cerca de 0,05); espécie do México\nB- Fêmur posterior com faixa escura basal larga (mais de 0,10); espécie do Panamá e da América do Sul")
    D60_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame77, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes homoeopus", frame77))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame77, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes terrens", frame77))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D61(frame):
    frame.pack_forget()
    frame78 = Frame(root, bg=cor1)
    frame78.pack(fill="both", expand=True)
    display_logo(frame78)

    # Pergunta 61 da Chave D
    D61_label = Label(frame78, text="Especie Aedes - Pergunta 61", font=("Times 20"), bg=cor1, fg="white")
    D61_label.pack(pady=10)
    D61_text = Text(frame78, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D61_text.insert(INSERT,
                    "A- Escudo (Ese) com parte clara anterior, transversa\nB- Escudo (Esc) sem parte clara anterior, transversa")
    D61_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame78, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D62(frame78))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame78, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D63(frame78))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D62(frame):
    frame.pack_forget()
    frame79 = Frame(root, bg=cor1)
    frame79.pack(fill="both", expand=True)
    display_logo(frame79)

    # Pergunta 62 da Chave D
    D62_label = Label(frame79, text="Especie Aedes - Pergunta 62", font=("Times 20"), bg=cor1, fg="white")
    D62_label.pack(pady=10)
    D62_text = Text(frame79, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D62_text.insert(INSERT, "A- Pós-pronoto (Ppn), com escamas claras\nB- Pós-pronoto (Ppn), com escamas escuras")
    D62_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame79, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes bertrami", frame79))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame79, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes aitkeni", frame79))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D63(frame):
    frame.pack_forget()
    frame80 = Frame(root, bg=cor1)
    frame80.pack(fill="both", expand=True)
    display_logo(frame80)

    # Pergunta 63 da Chave D
    D63_label = Label(frame80, text="Especie Aedes - Pergunta 63", font=("Times 20"), bg=cor1, fg="white")
    D63_label.pack(pady=10)
    D63_text = Text(frame80, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D63_text.insert(INSERT, "A- Tarsômero posterior Ta-5, claro\nB- Tarsômero posterior Ta -5, escuro")
    D63_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame80, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D64(frame80))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame80, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D65(frame80))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D64(frame):
    frame.pack_forget()
    frame81 = Frame(root, bg=cor1)
    frame81.pack(fill="both", expand=True)
    display_logo(frame81)

    # Pergunta 64 da Chave D
    D64_label = Label(frame81, text="Especie Aedes - Pergunta 64", font=("Times 20"), bg=cor1, fg="white")
    D64_label.pack(pady=10)
    D64_text = Text(frame81, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D64_text.insert(INSERT,
                    "A- Tarsômero anterior Ta-1 com cerca da metade apical, clara\nB- Tarsômero anterior Ta-1 com, sensivelmente, menos da metade apical, clara")
    D64_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame81, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes diazi", frame81))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame81, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes alboapicus", frame81))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D65(frame):
    frame.pack_forget()
    frame82 = Frame(root, bg=cor1)
    frame82.pack(fill="both", expand=True)
    display_logo(frame82)

    # Pergunta 65 da Chave D
    D65_label = Label(frame82, text="Especie Aedes - Pergunta 65", font=("Times 20"), bg=cor1, fg="white")
    D65_label.pack(pady=10)
    D65_text = Text(frame82, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D65_text.insert(INSERT,
                    "A- Mácula da área supra-alar (ASa) alcançando amplamente a sutura escudo-escutelo (see)\nB- Mácula da área supra-alar (ASa), raramente alcançando a sutura escudo-escutelo (see) e, quando isso ocorre, de maneira discreta")
    D65_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame82, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes casali", frame82))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame82, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes galindoi", frame82))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D66(frame):
    frame.pack_forget()
    frame83 = Frame(root, bg=cor1)
    frame83.pack(fill="both", expand=True)
    display_logo(frame83)

    # Pergunta 66 da Chave D
    D66_label = Label(frame83, text="Especie Aedes - Pergunta 66", font=("Times 20"), bg=cor1, fg="white")
    D66_label.pack(pady=10)
    D66_text = Text(frame83, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D66_text.insert(INSERT, "A- Cerdas acrosticais (SAc), ausentes\nB- Cerdas acrosticais (SAc), presentes")
    D66_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame83, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D67(frame83))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame83, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D69(frame83))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D67(frame):
    frame.pack_forget()
    frame84 = Frame(root, bg=cor1)
    frame84.pack(fill="both", expand=True)
    display_logo(frame84)

    # Pergunta 67 da Chave D
    D67_label = Label(frame84, text="Especie Aedes - Pergunta 67", font=("Times 20"), bg=cor1, fg="white")
    D67_label.pack(pady=10)
    D67_text = Text(frame84, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D67_text.insert(INSERT,
                    "A- Probóscide (P) mais curta do que o fêmur anterior\nB- Probóscide (P) mais longa do que o fêmur anterior")
    D67_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame84, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes podographicus", frame84))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame84, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D68(frame84))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D68(frame):
    frame.pack_forget()
    frame85 = Frame(root, bg=cor1)
    frame85.pack(fill="both", expand=True)
    display_logo(frame85)

    # Pergunta 68 da Chave D
    D68_label = Label(frame85, text="Especie Aedes - Pergunta 68", font=("Times 20"), bg=cor1, fg="white")
    D68_label.pack(pady=10)
    D68_text = Text(frame85, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D68_text.insert(INSERT,
                    "A- Vértice (V) com todas as escamas deitadas, claras\nB- Vértice (V) com as escamas deitadas, claras e escuras, estas tendendo a formar mancha submediana")
    D68_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame85, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes tehuantepec", frame85))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame85, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes schroederi", frame85))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D69(frame):
    frame.pack_forget()
    frame86 = Frame(root, bg=cor1)
    frame86.pack(fill="both", expand=True)
    display_logo(frame86)

    # Pergunta 69 da Chave D
    D69_label = Label(frame86, text="Especie Aedes - Pergunta 69", font=("Times 20"), bg=cor1, fg="white")
    D69_label.pack(pady=10)
    D69_text = Text(frame86, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D69_text.insert(INSERT,
                    "A- Fêmur mediano com a articulação com a tíbia correspondente, amplamente marcada\nB- Fêmur mediano com a articulação com a tíbia correspondente, sem marcação ou, quando existente, de maneira discreta")
    D69_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame86, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D70(frame86))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame86, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D76(frame86))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D70(frame):
    frame.pack_forget()
    frame87 = Frame(root, bg=cor1)
    frame87.pack(fill="both", expand=True)
    display_logo(frame87)

    # Pergunta 70 da Chave D
    D70_label = Label(frame87, text="Especie Aedes - Pergunta 70", font=("Times 20"), bg=cor1, fg="white")
    D70_label.pack(pady=10)
    D70_text = Text(frame87, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D70_text.insert(INSERT,
                    "A- Escamas subespiraculares (EASe), ausentes, se bem que é possível observar a presença de raros (um ou dois) desses elementos\nB- Escamas subespiraculares (EASe), presentes")
    D70_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame87, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes vargasi", frame87))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame87, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D71(frame87))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D71(frame):
    frame.pack_forget()
    frame88 = Frame(root, bg=cor1)
    frame88.pack(fill="both", expand=True)
    display_logo(frame88)

    # Pergunta 71 da Chave D
    D71_label = Label(frame88, text="Especie Aedes - Pergunta 71", font=("Times 20"), bg=cor1, fg="white")
    D71_label.pack(pady=10)
    D71_text = Text(frame88, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D71_text.insert(INSERT,
                    "A- Probóscide (P) mais curta ou de comprimento subigual ao do fêmur anterior\nB- Probóscide (P) mais longa do que o fêmur anterior")
    D71_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame88, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D72(frame88))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame88, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D73(frame88))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D72(frame):
    frame.pack_forget()
    frame89 = Frame(root, bg=cor1)
    frame89.pack(fill="both", expand=True)
    display_logo(frame89)

    # Pergunta 72 da Chave D
    D72_label = Label(frame89, text="Especie Aedes - Pergunta 72", font=("Times 20"), bg=cor1, fg="white")
    D72_label.pack(pady=10)
    D72_text = Text(frame89, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D72_text.insert(INSERT,
                    "A- Cerdas acrosticais (SAc) ausentes ou, então, com algumas escamas claras (EAc), como que representando a linha acrostical\nB- Cerdas acrosticais (SAc), presentes (linha acrostical)")
    D72_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame89, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes sumidero", frame89))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame89, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes gabriel", frame89))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D73(frame):
    frame.pack_forget()
    frame90 = Frame(root, bg=cor1)
    frame90.pack(fill="both", expand=True)
    display_logo(frame90)

    # Pergunta 73 da Chave D
    D73_label = Label(frame90, text="Especie Aedes - Pergunta 73", font=("Times 20"), bg=cor1, fg="white")
    D73_label.pack(pady=10)
    D73_text = Text(frame90, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D73_text.insert(INSERT,
                    "A- Fêmur anterior marcado na articulação com a tíbia correspondente\nB- Fêmur anterior não marcado na articulação com a tíbia correspondente")
    D73_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame90, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes amabilis", frame90))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame90, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D74(frame90))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D74(frame):
    frame.pack_forget()
    frame91 = Frame(root, bg=cor1)
    frame91.pack(fill="both", expand=True)
    display_logo(frame91)

    # Pergunta 74 da Chave D
    D74_label = Label(frame91, text="Especie Aedes - Pergunta 74", font=("Times 20"), bg=cor1, fg="white")
    D74_label.pack(pady=10)
    D74_text = Text(frame91, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D74_text.insert(INSERT,
                    "A- Vértice (V) dotado de área com escamas deitadas, largas, adjacente à de escamas estreitas e curvas, na sutura coronária (se)\nB- Vértice (V) com todas as escamas deitadas, estreitas e curvas")
    D74_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame91, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes idanus", frame91))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame91, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D75(frame91))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D75(frame):
    frame.pack_forget()
    frame92 = Frame(root, bg=cor1)
    frame92.pack(fill="both", expand=True)
    display_logo(frame92)

    # Pergunta 75 da Chave D
    D75_label = Label(frame92, text="Especie Aedes - Pergunta 75", font=("Times 20"), bg=cor1, fg="white")
    D75_label.pack(pady=10)
    D75_text = Text(frame92, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D75_text.insert(INSERT,
                    "A- Lobo mediano do escutelo (Eto), com escamas claras\nB- Lobo mediano do escutelo (Eto) sem escamas claras")
    D75_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame92, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes homoeopus", frame92))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame92, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes heteropus", frame92))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D76(frame):
    frame.pack_forget()
    frame93 = Frame(root, bg=cor1)
    frame93.pack(fill="both", expand=True)
    display_logo(frame93)

    # Pergunta 76 da Chave D
    D76_label = Label(frame93, text="Especie Aedes - Pergunta 76", font=("Times 20"), bg=cor1, fg="white")
    D76_label.pack(pady=10)
    D76_text = Text(frame93, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D76_text.insert(INSERT,
                    "A- Mácula da área supra-alar (ASa), alcançando a sutura escudo-escutelo (see)\nB- Mácula da área supra-alar (ASa), não alcançando a sutura escudo-escutelo")
    D76_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame93, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes daryi", frame93))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame93, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes galindoi", frame93))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D77(frame):
    frame.pack_forget()
    frame94 = Frame(root, bg=cor1)
    frame94.pack(fill="both", expand=True)
    display_logo(frame94)

    # Pergunta 77 da Chave D
    D77_label = Label(frame94, text="Especie Aedes - Pergunta 77", font=("Times 20"), bg=cor1, fg="white")
    D77_label.pack(pady=10)
    D77_text = Text(frame94, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D77_text.insert(INSERT, "A- Paratergito (Pa) sem escamas\nB- Paratergito (Pa) com escamas")
    D77_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame94, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D78(frame94))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame94, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D88(frame94))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D78(frame):
    frame.pack_forget()
    frame95 = Frame(root, bg=cor1)
    frame95.pack(fill="both", expand=True)
    display_logo(frame95)

    # Pergunta 78 da Chave D
    D78_label = Label(frame95, text="Especie Aedes - Pergunta 78", font=("Times 20"), bg=cor1, fg="white")
    D78_label.pack(pady=10)
    D78_text = Text(frame95, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D78_text.insert(INSERT,
                    "A- Escudo (Esc) coberto de escamas escuras ou bronzeadas\nB- Escudo (Esc) coberto de escamas claras, ao menos lateralmente")
    D78_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame95, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes calumnior", frame95))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame95, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D79(frame95))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D79(frame):
    frame.pack_forget()
    frame96 = Frame(root, bg=cor1)
    frame96.pack(fill="both", expand=True)
    display_logo(frame96)

    # Pergunta 79 da Chave D
    D79_label = Label(frame96, text="Especie Aedes - Pergunta 79", font=("Times 20"), bg=cor1, fg="white")
    D79_label.pack(pady=10)
    D79_text = Text(frame96, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D79_text.insert(INSERT,
                    "A- Escudo (Esc) com escamas douradas, em sua totalidade\nB- Escudo (Esc) com escamas claras dispostas nas margens laterais da área")
    D79_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame96, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D80(frame96))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame96, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D82(frame96))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D80(frame):
    frame.pack_forget()
    frame97 = Frame(root, bg=cor1)
    frame97.pack(fill="both", expand=True)
    display_logo(frame97)

    # Pergunta 80 da Chave D
    D80_label = Label(frame97, text="Especie Aedes - Pergunta 80", font=("Times 20"), bg=cor1, fg="white")
    D80_label.pack(pady=10)
    D80_text = Text(frame97, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D80_text.insert(INSERT,
                    "A- Pleura com ampla mancha clara, de escamas metálicas, contínua e estendendo-se desde o botão pré-alar (BPal), cobrindo todo o mesocatepisterno (Mcs)\nB- Pleura sem essa mancha clara contínua, mas subdividida em outras, menores, em número de três")
    D80_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame97, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes aurites", frame97))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame97, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D81(frame97))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D81(frame):
    frame.pack_forget()
    frame98 = Frame(root, bg=cor1)
    frame98.pack(fill="both", expand=True)
    display_logo(frame98)

    # Pergunta 81 da Chave D
    D81_label = Label(frame98, text="Especie Aedes - Pergunta 81", font=("Times 20"), bg=cor1, fg="white")
    D81_label.pack(pady=10)
    D81_text = Text(frame98, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D81_text.insert(INSERT,
                    "A- Área antealar (AAn), com escamas douradas\nB- Área antealar (AAn), com numerosas escamas bronzeadas")
    D81_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame98, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes grabhami", frame98))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame98, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes inaequalis", frame98))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D82(frame):
    frame.pack_forget()
    frame99 = Frame(root, bg=cor1)
    frame99.pack(fill="both", expand=True)
    display_logo(frame99)

    # Pergunta 82 da Chave D
    D82_label = Label(frame99, text="Especie Aedes - Pergunta 82", font=("Times 20"), bg=cor1, fg="white")
    D82_label.pack(pady=10)
    D82_text = Text(frame99, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D82_text.insert(INSERT,
                    "A- Área subespiracular (Ase), sem escamas\nB- Área subespiracular (Ase), com escamas (EAse)")
    D82_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame99, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D83(frame99))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame99, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D84(frame99))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D83(frame):
    frame.pack_forget()
    frame100 = Frame(root, bg=cor1)
    frame100.pack(fill="both", expand=True)
    display_logo(frame100)

    # Pergunta 83 da Chave D
    D83_label = Label(frame100, text="Especie Aedes - Pergunta 83", font=("Times 20"), bg=cor1, fg="white")
    D83_label.pack(pady=10)
    D83_text = Text(frame100, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D83_text.insert(INSERT,
                    "A- Área dorsocentral (AD), com linha contínua\nB- Área dorsocentral (AD), com linha tracejada")
    D83_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame100, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes walkeri", frame100))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame100, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes argyrites", frame100))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D84(frame):
    frame.pack_forget()
    frame101 = Frame(root, bg=cor1)
    frame101.pack(fill="both", expand=True)
    display_logo(frame101)

    # Pergunta 84 da Chave D
    D84_label = Label(frame101, text="Especie Aedes - Pergunta 84", font=("Times 20"), bg=cor1, fg="white")
    D84_label.pack(pady=10)
    D84_text = Text(frame101, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D84_text.insert(INSERT,
                    "A- Cerdas acrosticais (SAc) presentes; palpos maxilares (PlpM), escuros em ambos os sexos\nB- Cerdas acrosticais (SAc) ausentes; palpos maxilares (PlpM) com marcação clara apical na fêmea e basal em alguns dos palpômeros do macho")
    D84_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame101, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes busckii", frame101))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame101, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D85(frame101))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D85(frame):
    frame.pack_forget()
    frame102 = Frame(root, bg=cor1)
    frame102.pack(fill="both", expand=True)
    display_logo(frame102)

    # Pergunta 85 da Chave D
    D85_label = Label(frame102, text="Especie Aedes - Pergunta 85", font=("Times 20"), bg=cor1, fg="white")
    D85_label.pack(pady=10)
    D85_text = Text(frame102, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D85_text.insert(INSERT,
                    "A- Tarsômero posterior Ta-5, com escamas claras\nB- Tarsômero posterior Ta-5, somente com escamas escuras")
    D85_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame102, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D86(frame102))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame102, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D87(frame102))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D86(frame):
    frame.pack_forget()
    frame103 = Frame(root, bg=cor1)
    frame103.pack(fill="both", expand=True)
    display_logo(frame103)

    # Pergunta 86 da Chave D
    D86_label = Label(frame103, text="Especie Aedes - Pergunta 86", font=("Times 20"), bg=cor1, fg="white")
    D86_label.pack(pady=10)
    D86_text = Text(frame103, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D86_text.insert(INSERT,
                    "A- Tarsômero posterior Ta-5, inteiramente claro\nB- Tarsômero posterior Ta-5, com algumas escamas escuras no ápice e na face posterior")
    D86_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame103, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes ioliota", frame103))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame103, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes cozumelensis", frame103))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D87(frame):
    frame.pack_forget()
    frame104 = Frame(root, bg=cor1)
    frame104.pack(fill="both", expand=True)
    display_logo(frame104)

    # Pergunta 87 da Chave D
    D87_label = Label(frame104, text="Especie Aedes - Pergunta 87", font=("Times 20"), bg=cor1, fg="white")
    D87_label.pack(pady=10)
    D87_text = Text(frame104, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D87_text.insert(INSERT,
                    "A- Tarsômero posterior Ta-3, com claro basal até 0,5\nB- Tarsômero posterior Ta-3, com, pelo menos, 0,7 de claro basal")
    D87_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame104, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes bahamensis", frame104))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame104, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes albonotatus", frame104))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D88(frame):
    frame.pack_forget()
    frame105 = Frame(root, bg=cor1)
    frame105.pack(fill="both", expand=True)
    display_logo(frame105)

    # Pergunta 88 da Chave D
    D88_label = Label(frame105, text="Especie Aedes - Pergunta 88", font=("Times 20"), bg=cor1, fg="white")
    D88_label.pack(pady=10)
    D88_text = Text(frame105, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D88_text.insert(INSERT,
                    "A- Escudo (Esc) com a área acrostical (AA) como linha, sempre escura\nB- Escudo (Esc) com a área acrostical (AA) como linha clara, amarelada ou dourada")
    D88_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame105, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D89(frame105))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame105, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D103(frame105))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D89(frame):
    frame.pack_forget()
    frame106 = Frame(root, bg=cor1)
    frame106.pack(fill="both", expand=True)
    display_logo(frame106)

    # Pergunta 89 da Chave D
    D89_label = Label(frame106, text="Especie Aedes - Pergunta 89", font=("Times 20"), bg=cor1, fg="white")
    D89_label.pack(pady=10)
    D89_text = Text(frame106, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D89_text.insert(INSERT,
                    "A- Fêmur mediano sem indícios de escamas escuras\nB- Fêmur mediano com escamas claras evidentes")
    D89_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame106, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D90(frame106))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame106, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D92(frame106))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D90(frame):
    frame.pack_forget()
    frame107 = Frame(root, bg=cor1)
    frame107.pack(fill="both", expand=True)
    display_logo(frame107)

    # Pergunta 90 da Chave D
    D90_label = Label(frame107, text="Especie Aedes - Pergunta 90", font=("Times 20"), bg=cor1, fg="white")
    D90_label.pack(pady=10)
    D90_text = Text(frame107, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D90_text.insert(INSERT,
                    "A- Área dorsocentral (AD), como linha que não chega a atingir, anteriormente, a margem lateral do escudo (Esc)\nB- Área dorsocentral (AD), como linha que chega a atingir, anteriormente, a margem lateral do escudo (Esc)")
    D90_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame107, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes brevis", frame107))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame107, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D91(frame107))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D91(frame):
    frame.pack_forget()
    frame108 = Frame(root, bg=cor1)
    frame108.pack(fill="both", expand=True)
    display_logo(frame108)

    # Pergunta 91 da Chave D
    D91_label = Label(frame108, text="Especie Aedes - Pergunta 91", font=("Times 20"), bg=cor1, fg="white")
    D91_label.pack(pady=10)
    D91_text = Text(frame108, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D91_text.insert(INSERT,
                    "A- Escamas pós-pronotais (EPpn), de aspecto predominantemente estreito e de tonalidade amarelada; primeiro flagelômero (Flm-1), com escamas escuras\nB- Escamas pós-pronotais (EPpn), de aspecto predominantemente largo e de tonalidade branca; primeiro flagelômero (Flm-1) com escamas brancas")
    D91_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame108, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes quadrivittatus", frame108))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame108, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes lorraineae", frame108))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D92(frame):
    frame.pack_forget()
    frame109 = Frame(root, bg=cor1)
    frame109.pack(fill="both", expand=True)
    display_logo(frame109)

    # Pergunta 92 da Chave D
    D92_label = Label(frame109, text="Especie Aedes - Pergunta 92", font=("Times 20"), bg=cor1, fg="white")
    D92_label.pack(pady=10)
    D92_text = Text(frame109, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D92_text.insert(INSERT,
                    "A- Área antealar (AAn) com numerosas escamas bronzeadas (EAn), inferiormente situadas em relação às situadas na área supra-alar (ASa), estas dispostas em linha estreita e de tonalidade dourada\nB- Área antealar (AAn) com escamas esbranquiçadas, amareladas ou douradas (EAn), formando conjunto com as supra-alares (ESa)")
    D92_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame109, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D93(frame109))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame109, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D95(frame109))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D93(frame):
    frame.pack_forget()
    frame110 = Frame(root, bg=cor1)
    frame110.pack(fill="both", expand=True)
    display_logo(frame110)

    # Pergunta 93 da Chave D
    D93_label = Label(frame110, text="Especie Aedes - Pergunta 93", font=("Times 20"), bg=cor1, fg="white")
    D93_label.pack(pady=10)
    D93_text = Text(frame110, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D93_text.insert(INSERT,
                    "A- Área dorsocentral (AD) como linha alcançando a margem anterior do escudo (Esc); o toro (To), com pequeno conjunto mediano de escamas claras\nB- Área dorsocentral (AD) como linha que não chega a alcançar a margem anterior do escudo (Esc); o toro (To), com amplo anel de escamas claras, apenas dorsalmente interrompido")
    D93_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame110, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes sexlineatus", frame110))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame110, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D94(frame110))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D94(frame):
    frame.pack_forget()
    frame111 = Frame(root, bg=cor1)
    frame111.pack(fill="both", expand=True)
    display_logo(frame111)

    # Pergunta 94 da Chave D
    D94_label = Label(frame111, text="Especie Aedes - Pergunta 94", font=("Times 20"), bg=cor1, fg="white")
    D94_label.pack(pady=10)
    D94_text = Text(frame111, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D94_text.insert(INSERT,
                    "A- Tarsômero posterior Ta-3 com marcação clara; linha da área dorsocentral (AD) com a extremidade anterior terminando em mácula de escamas claras\nB- Tarsômero posterior Ta-3 sem marcação clara; linha da área dorsocentral (AD) sem mácula anterior de escamas claras")
    D94_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame111, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes whitmorei", frame111))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame111, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes pseudodominicii", frame111))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D95(frame):
    frame.pack_forget()
    frame112 = Frame(root, bg=cor1)
    frame112.pack(fill="both", expand=True)
    display_logo(frame112)

    # Pergunta 95 da Chave D
    D95_label = Label(frame112, text="Especie Aedes - Pergunta 95", font=("Times 20"), bg=cor1, fg="white")
    D95_label.pack(pady=10)
    D95_text = Text(frame112, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D95_text.insert(INSERT, "A- Base da veia alar R, com escamas claras\nB- Base da veia alar R, sem escamas claras")
    D95_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame112, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D96(frame112))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame112, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D98(frame112))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D96(frame):
    frame.pack_forget()
    frame113 = Frame(root, bg=cor1)
    frame113.pack(fill="both", expand=True)
    display_logo(frame113)

    # Pergunta 96 da Chave D
    D96_label = Label(frame113, text="Especie Aedes - Pergunta 96", font=("Times 20"), bg=cor1, fg="white")
    D96_label.pack(pady=10)
    D96_text = Text(frame113, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D96_text.insert(INSERT,
                    "A- Somente o tarsômero posterior Ta-1, dotado de marcação clara; haltere com o capitelo (Ca), dotado somente de escamas claras\nB- Os tarsômeros posteriores Ta-1 a 3, com marcações claras; haltere com o capitelo (Ca), dotado de escamas claras e escuras")
    D96_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame113, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes vanemdeni", frame113))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame113, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D97(frame113))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D97(frame):
    frame.pack_forget()
    frame114 = Frame(root, bg=cor1)
    frame114.pack(fill="both", expand=True)
    display_logo(frame114)

    # Pergunta 97 da Chave D
    D97_label = Label(frame114, text="Especie Aedes - Pergunta 97", font=("Times 20"), bg=cor1, fg="white")
    D97_label.pack(pady=10)
    D97_text = Text(frame114, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D97_text.insert(INSERT,
                    "A- Área dorsocentral (AD), como linha alcançando a margem anterior do escudo (Esc)\nB- Área dorsocentral (AD), como linha que não chega a atingir a margem anterior do escudo (Esc)")
    D97_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame114, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes aurivittatus", frame114))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame114, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes martinezi", frame114))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D98(frame):
    frame.pack_forget()
    frame115 = Frame(root, bg=cor1)
    frame115.pack(fill="both", expand=True)
    display_logo(frame115)

    # Pergunta 98 da Chave D
    D98_label = Label(frame115, text="Especie Aedes - Pergunta 98", font=("Times 20"), bg=cor1, fg="white")
    D98_label.pack(pady=10)
    D98_text = Text(frame115, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D98_text.insert(INSERT,
                    "A- Marcação clara basal dos tarsômeros posteriores, presente apenas nos Ta-1 e Ta-3\nB- Marcação clara basal dos tarsômeros posteriores, apenas em um dos segmentos ou em vários deles até Ta-4")
    D98_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame115, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes eleanorae", frame115))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame115, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D99(frame115))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D99(frame):
    frame.pack_forget()
    frame116 = Frame(root, bg=cor1)
    frame116.pack(fill="both", expand=True)
    display_logo(frame116)

    # Pergunta 99 da Chave D
    D99_label = Label(frame116, text="Especie Aedes - Pergunta 99", font=("Times 20"), bg=cor1, fg="white")
    D99_label.pack(pady=10)
    D99_text = Text(frame116, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D99_text.insert(INSERT,
                    "A- Marcação clara basal dos tarsômeros posteriores, variando de acordo com os segmentos, mas nunca em Ta-4\nB- Marcação clara basal dos tarsômeros posteriores, presente em Ta-1 a 4")
    D99_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame116, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes ecuadoriensis", frame116))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame116, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D100(frame116))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D100(frame):
    frame.pack_forget()
    frame117 = Frame(root, bg=cor1)
    frame117.pack(fill="both", expand=True)
    display_logo(frame117)

    # Pergunta 100 da Chave D
    D100_label = Label(frame117, text="Especie Aedes - Pergunta 100", font=("Times 20"), bg=cor1, fg="white")
    D100_label.pack(pady=10)
    D100_text = Text(frame117, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D100_text.insert(INSERT,
                     "A- Fêmur posterior com mácula clara distinta na parte escura da superfície anterior dos segmentos, sem confluir com a faixa basal clara\nB- Fêmur posterior sem mácula clara na parte escura da superfície anterior")
    D100_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame117, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D101(frame117))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame117, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes marinkellei", frame117))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D101(frame):
    frame.pack_forget()
    frame118 = Frame(root, bg=cor1)
    frame118.pack(fill="both", expand=True)
    display_logo(frame118)

    # Pergunta 101 da Chave D
    D101_label = Label(frame118, text="Especie Aedes - Pergunta 101", font=("Times 20"), bg=cor1, fg="white")
    D101_label.pack(pady=10)
    D101_text = Text(frame118, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D101_text.insert(INSERT,
                     "A- Marcação clara basal dos tarsômeros posteriores, presente apenas em Ta-1 e Ta-2\nB- Marcação clara basal dos tarsômeros posteriores, presente nos Ta-1 a 3")
    D101_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame118, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes guatemala", frame118))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame118, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D102(frame118))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D102(frame):
    frame.pack_forget()
    frame119 = Frame(root, bg=cor1)
    frame119.pack(fill="both", expand=True)
    display_logo(frame119)

    # Pergunta 102 da Chave D
    D102_label = Label(frame119, text="Especie Aedes - Pergunta 102", font=("Times 20"), bg=cor1, fg="white")
    D102_label.pack(pady=10)
    D102_text = Text(frame119, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D102_text.insert(INSERT,
                     "A- Escamas pós-espiraculares (EAP), presentes\nB- Escamas pós-espiraculares (EAP), ausentes")
    D102_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame119, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes guerrero", frame119))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame119, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes allotecnon", frame119))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D103(frame):
    frame.pack_forget()
    frame120 = Frame(root, bg=cor1)
    frame120.pack(fill="both", expand=True)
    display_logo(frame120)

    # Pergunta 103 da Chave D
    D103_label = Label(frame120, text="Especie Aedes - Pergunta 103", font=("Times 20"), bg=cor1, fg="white")
    D103_label.pack(pady=10)
    D103_text = Text(frame120, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D103_text.insert(INSERT,
                     "A- Escudo (Esc) com a parte anterior coberta de escamas douradas formando assim extensa mancha que se estende até a raiz da asa, enquanto na parte posterior pode-se observar a existência de escamas escuras\nB- Escudo (Esc) com a parte anterior não inteiramente dourada, com máculas ou faixas de escamas bronzeadas")
    D103_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame120, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes fulvithorax", frame120))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame120, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_D104(frame120))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_D104(frame):
    frame.pack_forget()
    frame121 = Frame(root, bg=cor1)
    frame121.pack(fill="both", expand=True)
    display_logo(frame121)

    # Pergunta 104 da Chave D
    D104_label = Label(frame121, text="Especie Aedes - Pergunta 104", font=("Times 20"), bg=cor1, fg="white")
    D104_label.pack(pady=10)
    D104_text = Text(frame121, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D104_text.insert(INSERT,
                     "A- Escudo (Esc) com somente a área acrostical (AA), como linha de aspecto dourado evidente na parte anterior\nB- Escudo (Esc) com não apenas a área acrostical (AA), mas também a área dorsocentral (AD), apresentando-se como linhas")
    D104_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame121, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes aurolineatus", frame121))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame121, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Aedes septemstriatus", frame121))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def create_chave_E_widgets():
    frame122 = Frame(root, bg=cor1)
    frame122.pack(fill="both", expand=True)
    display_logo(frame122)

    # Pergunta 1 da Chave E
    D1_label = Label(frame122, text="Especie Haemagogus - Pergunta 1", font=("Times 20"), bg=cor1, fg="white")
    D1_label.pack(pady=10)
    D1_text = Text(frame122, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    D1_text.insert(INSERT,
                   "A - Escamas do mesonoto marrom-escuras, com escamas prateadas em manchas ou linhas conspícuas nas áreas acrostical, antealar e pré-escutelar; pleura com três arcos verticais de escamas prateadas; mesonoto com linha acrostical de largura variável e se estendendo além da metade anterior> 2\n"
                   "B - Escamas do mesonoto metálicas (verde, azul, cobre ou bronze), com escamas prateadas apenas na área antealar; pleura com apenas uma faixa vertical larga de escamas prateadas. > 3")
    D1_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame122, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_E2(frame122))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame122, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_E3(frame122))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_E2(frame):
    frame.pack_forget()
    frame123 = Frame(root, bg=cor1)
    frame123.pack(fill="both", expand=True)
    display_logo(frame123)

    # Pergunta 2 da Chave E
    E2_label = Label(frame123, text="Especie Haemagogus - Pergunta 2", font=("Times 20"), bg=cor1, fg="white")
    E2_label.pack(pady=10)
    E2_text = Text(frame123, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    E2_text.insert(INSERT,
                   "A- Mancha de escamas mesepimeral com extensão moderadamente longa na parte superior; escamas eretas da cabeça completamente marrom-escuras ou enegrecidas.\nB- Mancha de escamas mesepimeral simples ou com uma extensão curta na parte superior; escamas eretas da cabeça completamente marrom-claras ou parcialmente brancas.")
    E2_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame123, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Haemagogus leucophoebus", frame123))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame123, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Haemagogus leucocelaenus", frame123))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_E3(frame):
    frame.pack_forget()
    frame124 = Frame(root, bg=cor1)
    frame124.pack(fill="both", expand=True)
    display_logo(frame124)

    # Pergunta 3 da Chave E
    E3_label = Label(frame124, text="Especie Haemagogus - Pergunta 3", font=("Times 20"), bg=cor1, fg="white")
    E3_label.pack(pady=10)
    E3_text = Text(frame124, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    E3_text.insert(INSERT,
                   "A- Cerdas no mesoposnoto, garras das pernas das fêmeas sem dente sub-basal.\nB- Sem cerdas no mesoposnoto, garras das pernas anteriores e médias das fêmeas com dente sub-basal.")
    E3_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame124, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Haemagogus celeste", frame124))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame124, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_E4(frame124))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_E4(frame):
    frame.pack_forget()
    frame125 = Frame(root, bg=cor1)
    frame125.pack(fill="both", expand=True)
    display_logo(frame125)

    # Pergunta 4 da Chave E
    E4_label = Label(frame125, text="Especie Haemagogus - Pergunta 4", font=("Times 20"), bg=cor1, fg="white")
    E4_label.pack(pady=10)
    E4_text = Text(frame125, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    E4_text.insert(INSERT, "A- Pós-pronoto sem escamas.\nB- Pós-pronoto com escamas.")
    E4_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame125, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Haemagogus tropicalis", frame125))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame125, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_E5(frame125))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_E5(frame):
    frame.pack_forget()
    frame126 = Frame(root, bg=cor1)
    frame126.pack(fill="both", expand=True)
    display_logo(frame126)

    # Pergunta 5 da Chave E
    E5_label = Label(frame126, text="Especie Haemagogus - Pergunta 5", font=("Times 20"), bg=cor1, fg="white")
    E5_label.pack(pady=10)
    E5_text = Text(frame126, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    E5_text.insert(INSERT,
                   "A- Tarsos médios cobertos somente com escamas escuras, probóscida curta, com menos de 1,40X do comprimento do fêmur anterior, espaço entre os olhos estreito (1-2 diâmetros de omatídeos), abdome com escamas prateadas dorsais restritas aos tergitos IV-VII ou ausentes, sem manchas nos joelhos nos fêmures médios e posteriores, mesocatepisterno inferior com cerda bem desenvolvida.\nB- Tudo igual acima, mas sem cerda bem desenvolvida no mesocatepisterno inferior.")
    E5_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame126, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_E6(frame126))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame126, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_E7(frame126))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_E6(frame):
    frame.pack_forget()
    frame127 = Frame(root, bg=cor1)
    frame127.pack(fill="both", expand=True)
    display_logo(frame127)

    # Pergunta 6 da Chave E
    E6_label = Label(frame127, text="Especie Haemagogus - Pergunta 6", font=("Times 20"), bg=cor1, fg="white")
    E6_label.pack(pady=10)
    E6_text = Text(frame127, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    E6_text.insert(INSERT,
                   "A- Fêmur posterior com escamas prateadas estendendo-se anteriormente quase até o ápice; escamas escuras do abdome usualmente predominantemente purpúreas, mas com escamas de matiz azul-esverdeado na margem distal dos tergitos V-VIII; probóscida com 1,10-1,15X o comprimento do fêmur anterior.\nB- Fêmur posterior com escamas prateadas não se estendendo anteriormente além de 0,75 do comprimento; escamas escuras dos tergitos abdominais I-VII purpúreas; probóscida com 1,25X o comprimento do fêmur anterior.")
    E6_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame127, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Haemagogus janthinomys", frame127))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame127, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Haemagogus albomaculatus", frame127))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_E7(frame):
    frame.pack_forget()
    frame128 = Frame(root, bg=cor1)
    frame128.pack(fill="both", expand=True)
    display_logo(frame128)

    # Pergunta 7 da Chave E
    E7_label = Label(frame128, text="Especie Haemagogus - Pergunta 7", font=("Times 20"), bg=cor1, fg="white")
    E7_label.pack(pady=10)
    E7_text = Text(frame128, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    E7_text.insert(INSERT,
                   "A- Probóscida da fêmea curta, 1,10 do comprimento do fêmur anterior; escamas do mesonoto predominantemente azul-escuras; coxas com manchas conspícuas de escamas escuras; R2+3 1,20-1,50 de R2, escamas abdominais escuras purpúreas.\nB- Probóscida da fêmea com pelo menos 1,20 do fêmur anterior; escamas do mesonoto predominantemente cúpreas ou bronzeadas claras; coxas totalmente prateadas ou raramente com escamas escuras na coxa média; R2+3 1,20-1,50 de R2, escamas abdominais escuras com nuanças de azul, verde ou dourado; lobos antepronotais totalmente cobertos com escamas prateadas; tergitos abdominais sem escamas prateadas dorsalmente; escamas pós-pronotais douradas a cúpreas, escamas decumbentes do vértice verdes a douradas.")
    E7_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame128, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Haemagogus baresi", frame128))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame128, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("Haemagogus spegazzinii", frame128))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def create_chave_F_widgets():
    frame129 = Frame(root, bg=cor1)
    frame129.pack(fill="both", expand=True)
    display_logo(frame129)

    # Pergunta 1 da Chave F
    F1_label = Label(frame129, text="Especie Psorophora - Pergunta 1", font=("Times 20"), bg=cor1, fg="white")
    F1_label.pack(pady=10)
    F1_text = Text(frame129, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F1_text.insert(INSERT,
                   "A - Escamas do mesonoto marrom-escuras, com escamas prateadas em manchas ou linhas conspícuas nas áreas acrostical, antealar e pré-escutelar; pleura com três arcos verticais de escamas prateadas; mesonoto com linha acrostical de largura variável e se estendendo além da metade anterior> 2\n"
                   "B - Escamas do mesonoto metálicas (verde, azul, cobre ou bronze), com escamas prateadas apenas na área antealar; pleura com apenas uma faixa vertical larga de escamas prateadas. > 3")
    F1_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame129, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F2(frame129))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame129, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F3(frame129))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F2(frame):
    frame.pack_forget()
    frame130 = Frame(root, bg=cor1)
    frame130.pack(fill="both", expand=True)
    display_logo(frame130)

    # Pergunta 2 da Chave F
    F2_label = Label(frame130, text="Especie Haemagogus - Pergunta 2", font=("Times 20"), bg=cor1, fg="white")
    F2_label.pack(pady=10)
    F2_text = Text(frame130, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F2_text.insert(INSERT,
                   "A- Tarsômeros escuros em todas as pernas, isto é, destituídos de marcação clara.\nB- Tarsômeros com marcação clara, pelo menos em algumas pernas.")
    F2_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame130, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F3(frame130))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame130, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F5(frame130))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F3(frame):
    frame.pack_forget()
    frame131 = Frame(root, bg=cor1)
    frame131.pack(fill="both", expand=True)
    display_logo(frame131)

    # Pergunta 3 da Chave F
    F3_label = Label(frame131, text="Especie Haemagogus - Pergunta 3", font=("Times 20"), bg=cor1, fg="white")
    F3_label.pack(pady=10)
    F3_text = Text(frame131, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F3_text.insert(INSERT,
                   "A- Fêmur posterior com marcação clara na articulação com a tíbia.\nB- Fêmur posterior sem marcação clara na articulação com a tíbia.")
    F3_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame131, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("saeva", frame131))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame131, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F4(frame131))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F4(frame):
    frame.pack_forget()
    frame132 = Frame(root, bg=cor1)
    frame132.pack(fill="both", expand=True)
    display_logo(frame132)

    # Pergunta 4 da Chave F
    F4_label = Label(frame132, text="Especie Haemagogus - Pergunta 4", font=("Times 20"), bg=cor1, fg="white")
    F4_label.pack(pady=10)
    F4_text = Text(frame132, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F4_text.insert(INSERT,
                   "A- Pleura com escamas abundantes e uniformemente esparsas.\nB- Pleura com escamas pouco numerosas e distribuídas em conjuntos.")
    F4_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame132, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("cilipes", frame132))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame132, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("lineata", frame132))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F5(frame):
    frame.pack_forget()
    frame133 = Frame(root, bg=cor1)
    frame133.pack(fill="both", expand=True)
    display_logo(frame133)

    # Pergunta 5 da Chave F
    F5_label = Label(frame133, text="Especie Haemagogus - Pergunta 5", font=("Times 20"), bg=cor1, fg="white")
    F5_label.pack(pady=10)
    F5_text = Text(frame133, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F5_text.insert(INSERT,
                   "A- Escudo com linha longitudinal mediana de escamas escuras.\nB- Escudo com linha longitudinal de escamas douradas.")
    F5_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame133, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("holmbergii", frame133))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame133, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F6(frame133))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F6(frame):
    frame.pack_forget()
    frame134 = Frame(root, bg=cor1)
    frame134.pack(fill="both", expand=True)
    display_logo(frame134)

    # Pergunta 6 da Chave F
    F6_label = Label(frame134, text="Especie Haemagogus - Pergunta 6", font=("Times 20"), bg=cor1, fg="white")
    F6_label.pack(pady=10)
    F6_text = Text(frame134, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F6_text.insert(INSERT, "A- Escamas da asa totalmente escuras.\nB- Escamas da asa claras e escuras.")
    F6_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame134, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("ciliata", frame134))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame134, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("pallescens", frame134))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F7(frame):
    frame.pack_forget()
    frame135 = Frame(root, bg=cor1)
    frame135.pack(fill="both", expand=True)
    display_logo(frame135)

    # Pergunta 7 da Chave F
    F7_label = Label(frame135, text="Especie Haemagogus - Pergunta 7", font=("Times 20"), bg=cor1, fg="white")
    F7_label.pack(pady=10)
    F7_text = Text(frame135, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F7_text.insert(INSERT, "A- Garras tarsais (G) femininas, denteadas.\nB- Garras tarsais (G) femininas, lisas.")
    F7_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame135, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F8(frame135))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame135, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F21(frame135))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F8(frame):
    frame.pack_forget()
    frame136 = Frame(root, bg=cor1)
    frame136.pack(fill="both", expand=True)
    display_logo(frame136)

    # Pergunta 8 da Chave F
    F8_label = Label(frame136, text="Especie Haemagogus - Pergunta 8", font=("Times 20"), bg=cor1, fg="white")
    F8_label.pack(pady=10)
    F8_text = Text(frame136, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F8_text.insert(INSERT,
                   "A- Tarsômeros inteiramente escuros, sem marcação clara.\nB- Tarsômeros com marcação clara, pelo menos em alguns segmentos proximais posteriores.")
    F8_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame136, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F9(frame136))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame136, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F11(frame136))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F9(frame):
    frame.pack_forget()
    frame137 = Frame(root, bg=cor1)
    frame137.pack(fill="both", expand=True)
    display_logo(frame137)

    # Pergunta 9 da Chave F
    F9_label = Label(frame137, text="Especie Haemagogus - Pergunta 9", font=("Times 20"), bg=cor1, fg="white")
    F9_label.pack(pady=10)
    F9_text = Text(frame137, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F9_text.insert(INSERT,
                   "A- Escudo revestido por escamas claras, de maneira esparsa.\nB- Escudo com escamas escuras na área central e posterior, além de elementos de tonalidade escura na área anterior e nas áreas laterais.")
    F9_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame137, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F10(frame137))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame137, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("lanei", frame137))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F10(frame):
    frame.pack_forget()
    frame138 = Frame(root, bg=cor1)
    frame138.pack(fill="both", expand=True)
    display_logo(frame138)

    # Pergunta 10 da Chave F
    F10_label = Label(frame138, text="Especie Haemagogus - Pergunta 10", font=("Times 20"), bg=cor1, fg="white")
    F10_label.pack(pady=10)
    F10_text = Text(frame138, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F10_text.insert(INSERT,
                    "A- Abdome com os tergitos dotados de áreas triangulares claras, em situação apicomediana.\nB- Abdome com os tergitos dotados de áreas claras em situação apicolateral.")
    F10_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame138, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("cyanescens", frame138))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame138, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("johnstonii", frame138))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F11(frame):
    frame.pack_forget()
    frame139 = Frame(root, bg=cor1)
    frame139.pack(fill="both", expand=True)
    display_logo(frame139)

    # Pergunta 11 da Chave F
    F11_label = Label(frame139, text="Especie Haemagogus - Pergunta 11", font=("Times 20"), bg=cor1, fg="white")
    F11_label.pack(pady=10)
    F11_text = Text(frame139, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F11_text.insert(INSERT,
                    "A- Tarsômero posterior Ta-5, inteiramente escuro.\nB- Tarsômero posterior Ta-5, inteiramente claro.")
    F11_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame139, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F12(frame139))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame139, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F14(frame139))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F12(frame):
    frame.pack_forget()
    frame140 = Frame(root, bg=cor1)
    frame140.pack(fill="both", expand=True)
    display_logo(frame140)

    # Pergunta 12 da Chave F
    F12_label = Label(frame140, text="Especie Haemagogus - Pergunta 12", font=("Times 20"), bg=cor1, fg="white")
    F12_label.pack(pady=10)
    F12_text = Text(frame140, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F12_text.insert(INSERT,
                    "A- Escudo com escamas escuras na área central e claras amareladas nas laterais.\nB- Escudo com escamas claras esparsas.")
    F12_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame140, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",command=lambda: set_especie("varipes",frame140))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame140, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F13(frame140))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F13(frame):
    frame.pack_forget()
    frame141 = Frame(root, bg=cor1)
    frame141.pack(fill="both", expand=True)
    display_logo(frame141)

    # Pergunta 13 da Chave F
    F13_label = Label(frame141, text="Especie Haemagogus - Pergunta 13", font=("Times 20"), bg=cor1, fg="white")
    F13_label.pack(pady=10)
    F13_text = Text(frame141, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F13_text.insert(INSERT,
                    "A- Escamas laterais do escudo de tonalidade amarelada ou dourada.\nB- Escamas laterais do escudo de tonalidade branca.")
    F13_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame141, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("discrucians", frame141))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame141, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("johnstoni", frame141))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F14(frame):
    frame.pack_forget()
    frame142 = Frame(root, bg=cor1)
    frame142.pack(fill="both", expand=True)
    display_logo(frame142)

    # Pergunta 14 da Chave F
    F14_label = Label(frame142, text="Especie Haemagogus - Pergunta 14", font=("Times 20"), bg=cor1, fg="white")
    F14_label.pack(pady=10)
    F14_text = Text(frame142, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F14_text.insert(INSERT,
                    "A- Escudo com escamas escuras na área central e claras na lateral.\nB- Escudo com escamas claras e escuras, misturadas de maneira um tanto homogênea.")
    F14_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame142, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F15(frame142))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame142, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F20(frame142))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F15(frame):
    frame.pack_forget()
    frame143 = Frame(root, bg=cor1)
    frame143.pack(fill="both", expand=True)
    display_logo(frame143)

    # Pergunta 15 da Chave F
    F15_label = Label(frame143, text="Especie Haemagogus - Pergunta 15", font=("Times 20"), bg=cor1, fg="white")
    F15_label.pack(pady=10)
    F15_text = Text(frame143, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F15_text.insert(INSERT,
                    "A- Fêmures sem marcação clara na articulação com as tíbias.\nB- Fêmures com marcação clara na articulação com as tíbias.")
    F15_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame143, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F16(frame143))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame143, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F17(frame143))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F16(frame):
    frame.pack_forget()
    frame144 = Frame(root, bg=cor1)
    frame144.pack(fill="both", expand=True)
    display_logo(frame144)

    # Pergunta 16 da Chave F
    F16_label = Label(frame144, text="Especie Haemagogus - Pergunta 16", font=("Times 20"), bg=cor1, fg="white")
    F16_label.pack(pady=10)
    F16_text = Text(frame144, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F16_text.insert(INSERT,
                    "A- Lobo mediano do escutelo (Eto) com escamas escuras.\nB- Lobo mediano do escutelo (Eto) com escamas claras.")
    F16_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame144, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("circumflava", frame144))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame144, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("totonaci", frame144))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F17(frame):
    frame.pack_forget()
    frame145 = Frame(root, bg=cor1)
    frame145.pack(fill="both", expand=True)
    display_logo(frame145)

    # Pergunta 17 da Chave F
    F17_label = Label(frame145, text="Especie Haemagogus - Pergunta 17", font=("Times 20"), bg=cor1, fg="white")
    F17_label.pack(pady=10)
    F17_text = Text(frame145, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F17_text.insert(INSERT,
                    "A- Área pré-escutelar (Apr) e escutelo (Eto) com escamas claras.\nB- Área pré-escutelar (Apr) e escutelo (Eto) com escamas escuras.")
    F17_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame145, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("champerico", frame145))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame145, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F18(frame145))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F18(frame):
    frame.pack_forget()
    frame146 = Frame(root, bg=cor1)
    frame146.pack(fill="both", expand=True)
    display_logo(frame146)

    # Pergunta 18 da Chave F
    F18_label = Label(frame146, text="Especie Haemagogus - Pergunta 18", font=("Times 20"), bg=cor1, fg="white")
    F18_label.pack(pady=10)
    F18_text = Text(frame146, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F18_text.insert(INSERT,
                    "A- Escudo com escamas laterais claras, esbranquiçadas.\nB- Escudo com escamas laterais claras, amareladas ou douradas.")
    F18_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame146, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("albipes", frame146))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame146, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F19(frame146))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F19(frame):
    frame.pack_forget()
    frame147 = Frame(root, bg=cor1)
    frame147.pack(fill="both", expand=True)
    display_logo(frame147)

    # Pergunta 19 da Chave F
    F19_label = Label(frame147, text="Especie Haemagogus - Pergunta 19", font=("Times 20"), bg=cor1, fg="white")
    F19_label.pack(pady=10)
    F19_text = Text(frame147, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F19_text.insert(INSERT,
                    "A- Escudo com as escamas claras laterais formando faixa bem definida.\nB- Escudo com as escamas claras laterais, sem formar faixa definida.")
    F19_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame147, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("lutzii", frame147))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame147, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("forceps", frame147))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F20(frame):
    frame.pack_forget()
    frame148 = Frame(root, bg=cor1)
    frame148.pack(fill="both", expand=True)
    display_logo(frame148)

    # Pergunta 20 da Chave F
    F20_label = Label(frame148, text="Especie Haemagogus - Pergunta 20", font=("Times 20"), bg=cor1, fg="white")
    F20_label.pack(pady=10)
    F20_text = Text(frame148, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F20_text.insert(INSERT,
                    "A- Escudo coberto de escamas predominantemente escuras.\nB- Escudo coberto de escamas claras e escuras esparsas e em número equivalente.")
    F20_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame148, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("melanota", frame148))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame148, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("ferox", frame148))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F21(frame):
    frame.pack_forget()
    frame149 = Frame(root, bg=cor1)
    frame149.pack(fill="both", expand=True)
    display_logo(frame149)

    # Pergunta 21 da Chave F
    F21_label = Label(frame149, text="Especie Haemagogus - Pergunta 21", font=("Times 20"), bg=cor1, fg="white")
    F21_label.pack(pady=10)
    F21_text = Text(frame149, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F21_text.insert(INSERT, "A- Asa com escamas totalmente escuras.\nB- Asa com, pelo menos, algumas escamas claras.")
    F21_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame149, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("cingulata", frame149))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame149, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F22(frame149))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F22(frame):
    frame.pack_forget()
    frame150 = Frame(root, bg=cor1)
    frame150.pack(fill="both", expand=True)
    display_logo(frame150)

    # Pergunta 22 da Chave F
    F22_label = Label(frame150, text="Especie Haemagogus - Pergunta 22", font=("Times 20"), bg=cor1, fg="white")
    F22_label.pack(pady=10)
    F22_text = Text(frame150, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F22_text.insert(INSERT,
                    "A- Escamas claras da asa limitadas à costa (C); abdome com os tergitos dotados de faixa clara apical.\nB- Escamas claras da asa em todas as veias; abdome com os tergitos dotados de áreas claras triangulares apicolaterais.")
    F22_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame150, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("dimidiata", frame150))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame150, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F23(frame150))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F23(frame):
    frame.pack_forget()
    frame151 = Frame(root, bg=cor1)
    frame151.pack(fill="both", expand=True)
    display_logo(frame151)

    # Pergunta 23 da Chave F
    F23_label = Label(frame151, text="Especie Haemagogus - Pergunta 23", font=("Times 20"), bg=cor1, fg="white")
    F23_label.pack(pady=10)
    F23_text = Text(frame151, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F23_text.insert(INSERT,
                    "A- Tarsômero Ta-1, de todas as pernas, com anel mediano de escamas claras, ou então, apresentando-se salpicado com tais elementos.\nB- Tarsômero Ta-1, de todas as pernas, escuro, podendo apenas existir marcação clara basal.")
    F23_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame151, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F24(frame151))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame151, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F27(frame151))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F24(frame):
    frame.pack_forget()
    frame152 = Frame(root, bg=cor1)
    frame152.pack(fill="both", expand=True)
    display_logo(frame152)

    # Pergunta 24 da Chave F
    F24_label = Label(frame152, text="Especie Haemagogus - Pergunta 24", font=("Times 20"), bg=cor1, fg="white")
    F24_label.pack(pady=10)
    F24_text = Text(frame152, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F24_text.insert(INSERT,
                    "A- Tarsômero Ta-1, de todas as pernas, com anel claro mediano.\nB- Tarsômero Ta-1, de todas as pernas, salpicado de escamas claras.")
    F24_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame152, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F25(frame152))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame152, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F26(frame152))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F25(frame):
    frame.pack_forget()
    frame153 = Frame(root, bg=cor1)
    frame153.pack(fill="both", expand=True)
    display_logo(frame153)

    # Pergunta 25 da Chave F
    F25_label = Label(frame153, text="Especie Haemagogus - Pergunta 25", font=("Times 20"), bg=cor1, fg="white")
    F25_label.pack(pady=10)
    F25_text = Text(frame153, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F25_text.insert(INSERT,
                    "A- Abdome com os tergitos predominantemente cobertos de escamas claras.\nB- Abdome com os tergitos predominantemente cobertos de escamas escuras.")
    F25_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame153, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("jamaicensis", frame153))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame153, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("confinnis", frame153))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F26(frame):
    frame.pack_forget()
    frame154 = Frame(root, bg=cor1)
    frame154.pack(fill="both", expand=True)
    display_logo(frame154)

    # Pergunta 26 da Chave F
    F26_label = Label(frame154, text="Especie Haemagogus - Pergunta 26", font=("Times 20"), bg=cor1, fg="white")
    F26_label.pack(pady=10)
    F26_text = Text(frame154, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F26_text.insert(INSERT, "A- Veia A I com extremidade distal clara.\nB- Veia A com extremidade distal escura.")
    F26_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame154, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("signipennis", frame154))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame154, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("discolor", frame154))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F27(frame):
    frame.pack_forget()
    frame155 = Frame(root, bg=cor1)
    frame155.pack(fill="both", expand=True)
    display_logo(frame155)

    # Pergunta 27 da Chave F
    F27_label = Label(frame155, text="Especie Haemagogus - Pergunta 27", font=("Times 20"), bg=cor1, fg="white")
    F27_label.pack(pady=10)
    F27_text = Text(frame155, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F27_text.insert(INSERT,
                    "A- Escamas alares escuras formando uma ou mais manchas.\nB- Escamas alares, escuras e claras, homogeneamente misturadas.")
    F27_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame155, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F28(frame155))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame155, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_F29(frame155))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F28(frame):
    frame.pack_forget()
    frame156 = Frame(root, bg=cor1)
    frame156.pack(fill="both", expand=True)
    display_logo(frame156)

    # Pergunta 28 da Chave F
    F28_label = Label(frame156, text="Especie Haemagogus - Pergunta 28", font=("Times 20"), bg=cor1, fg="white")
    F28_label.pack(pady=10)
    F28_text = Text(frame156, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F28_text.insert(INSERT,
                    "A- Veia R4 +5 com uma só mancha escura, nítida e localizada basalmente.\nB- Veia R4+5 com várias manchas escuras.")
    F28_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame156, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("varinervis", frame156))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame156, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("paulli", frame156))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_F29(frame):
    frame.pack_forget()
    frame157 = Frame(root, bg=cor1)
    frame157.pack(fill="both", expand=True)
    display_logo(frame157)

    # Pergunta 29 da Chave F
    F29_label = Label(frame157, text="Especie Haemagogus - Pergunta 29", font=("Times 20"), bg=cor1, fg="white")
    F29_label.pack(pady=10)
    F29_text = Text(frame157, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    F29_text.insert(INSERT,
                    "A- Escudo uniformemente revestido de escamas claras; probóscide ventralmente clara porém sem demarcação nítida.\nB- Escudo com áreas escuras e irregularmente revestido de escamas douradas ou bronzeadas; probóscide com área ventral clara bem demarcada, frequentemente formando anel distinto.")
    F29_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame157, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("insularis", frame157))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame157, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("pygmaea", frame157))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def create_chave_g_widgets():
    frame158 = Frame(root, bg=cor1)
    frame158.pack(fill="both", expand=True)
    display_logo(frame158)

    # Pergunta 1 da Chave G
    G1_label = Label(frame158, text="Gêneros Sabethini - Pergunta 1", font=("Times 20"), bg=cor1, fg="white")
    G1_label.pack(pady=10)
    G1_text = Text(frame158, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    G1_text.insert(INSERT,
                   "A - Cerdas pré-espiraculares (SPe) presentes. > G2\nB - Cerdas pré-espiraculares (SPe) ausentes, substituídas por escamas. > Gênero =  limatus | SALVAR")
    G1_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame158, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_G2(frame158))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame158, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Limatus", frame158))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_G2(frame):
    frame.pack_forget()
    frame159 = Frame(root, bg=cor1)
    frame159.pack(fill="both", expand=True)
    display_logo(frame159)

    # Pergunta 2 da Chave G
    G2_label = Label(frame159, text="Gêneros Sabethini - Pergunta 2", font=("Times 20"), bg=cor1, fg="white")
    G2_label.pack(pady=10)
    G2_text = Text(frame159, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    G2_text.insert(INSERT,
                   "A- Calíptera superior (Cs) com fileira marginal de cerdas.\n \nB- Calíptera superior (Cs) sem fileira de cerdas ou então com, no máximo, quatro elementos.")
    G2_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame159, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_G3(frame159))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame159, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_G4(frame159))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_G3(frame):
    frame.pack_forget()
    frame160 = Frame(root, bg=cor1)
    frame160.pack(fill="both", expand=True)
    display_logo(frame160)

    # Pergunta 3 da Chave G
    G3_label = Label(frame160, text="Gêneros Sabethini - Pergunta 3", font=("Times 20"), bg=cor1, fg="white")
    G3_label.pack(pady=10)
    G3_text = Text(frame160, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    G3_text.insert(INSERT,
                   "A- Tíbia posterior sem faixa clara no meio; área pós-procoxal (Mpx) sem escamas.\n \nB- Tíbia posterior com ampla faixa clara mediana; área pós-procoxal (Mpx) com escamas.")
    G3_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame160, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Trichoprosopon", frame160))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame160, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Shannoniana", frame160))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_G4(frame):
    frame.pack_forget()
    frame161 = Frame(root, bg=cor1)
    frame161.pack(fill="both", expand=True)
    display_logo(frame161)

    # Pergunta 4 da Chave G
    G4_label = Label(frame161, text="Gêneros Sabethini - Pergunta 4", font=("Times 20"), bg=cor1, fg="white")
    G4_label.pack(pady=10)
    G4_text = Text(frame161, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    G4_text.insert(INSERT,
                   "A- Probóscide sensivelmente mais longa do que o fêmur anterior.\n \nB- Probóscide subigual ou sensivelmente mais curta do que o fêmur anterior.")
    G4_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame161, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_G5(frame161))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame161, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_G9(frame161))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_G5(frame):
    frame.pack_forget()
    frame162 = Frame(root, bg=cor1)
    frame162.pack(fill="both", expand=True)
    display_logo(frame162)

    # Pergunta 5 da Chave G
    G5_label = Label(frame162, text="Gêneros Sabethini - Pergunta 5", font=("Times 20"), bg=cor1, fg="white")
    G5_label.pack(pady=10)
    G5_text = Text(frame162, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    G5_text.insert(INSERT,
                   "A- Escutelo (Eto) com escamas de reflexos prateados.\n \nB- Escutelo (Eto) com escamas escuras e sem reflexos prateados.")
    G5_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame162, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_G6(frame162))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame162, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_G7(frame162))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_G6(frame):
    frame.pack_forget()
    frame163 = Frame(root, bg=cor1)
    frame163.pack(fill="both", expand=True)
    display_logo(frame163)

    # Pergunta 6 da Chave G
    G6_label = Label(frame163, text="Gêneros Sabethini - Pergunta 6", font=("Times 20"), bg=cor1, fg="white")
    G6_label.pack(pady=10)
    G6_text = Text(frame163, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    G6_text.insert(INSERT,
                   "A- Probóscide sensivelmente mais longa do que o fêmur anterior e curva na extremidade distal.\n \nB- Probóscide pouco mais longa do que o fêmur anterior e dilatada na extremidade distal.")
    G6_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame163, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Phoniomyia", frame163))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame163, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Wyeomyia", frame163))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_G7(frame):
    frame.pack_forget()
    frame164 = Frame(root, bg=cor1)
    frame164.pack(fill="both", expand=True)
    display_logo(frame164)

    # Pergunta 7 da Chave G
    G7_label = Label(frame164, text="Gêneros Sabethini - Pergunta 7", font=("Times 20"), bg=cor1, fg="white")
    G7_label.pack(pady=10)
    G7_text = Text(frame164, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    G7_text.insert(INSERT,
                   "A- Escamas do occipício (Occ) com reflexos prateados.\n \nB- Escamas do occipício (Occ) sem reflexos prateados.")
    G7_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame164, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Johnbelkinia", frame164))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame164, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_G8(frame164))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_G8(frame):
    frame.pack_forget()
    frame165 = Frame(root, bg=cor1)
    frame165.pack(fill="both", expand=True)
    display_logo(frame165)

    # Pergunta 8 da Chave G
    G8_label = Label(frame165, text="Gêneros Sabethini - Pergunta 8", font=("Times 20"), bg=cor1, fg="white")
    G8_label.pack(pady=10)
    G8_text = Text(frame165, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    G8_text.insert(INSERT,
                   "A- Área pós-pronotal (Ppn) com uma ou mais cerdas posteriores.\n \nB- Área pós-pronotal (Ppn) sem cerdas posteriores.")
    G8_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame165, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Runchomyia", frame165))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame165, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Isostomyia", frame165))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_G9(frame):
    frame.pack_forget()
    frame166 = Frame(root, bg=cor1)
    frame166.pack(fill="both", expand=True)
    display_logo(frame166)

    # Pergunta 9 da Chave G
    G9_label = Label(frame166, text="Gêneros Sabethini - Pergunta 9", font=("Times 20"), bg=cor1, fg="white")
    G9_label.pack(pady=10)
    G9_text = Text(frame166, height=5, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    G9_text.insert(INSERT,
                   "A- Lobos antepronotais (Ap) bem desenvolvidos; mosquitos de aparência metálica.\n \nB- Lobos antepronotais (Ap) normalmente desenvolvidos; mosquitos escuros.")
    G9_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame166, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero_and_continue("Sabethes", frame166, create_chave_H_widgets))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame166, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Wyeomyia", frame166))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def create_chave_H_widgets():
    frame167 = Frame(root, bg=cor1)
    frame167.pack(fill="both", expand=True)
    display_logo(frame167)

    # Pergunta 1 da Chave H
    H1_label = Label(frame167, text="Especies Sabethes - Pergunta 1", font=("Times 20"), bg=cor1, fg="white")
    H1_label.pack(pady=10)
    H1_text = Text(frame167, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H1_text.insert(INSERT,
                   "A - - cerda pré-alar presente (subgênero Davismyia) > Especie =  petrocchiae \n \nB- Ccerda pré-alar ausente > H2")
    H1_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame167, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("petrocchiae", frame167))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame167, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H2(frame167))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H2(frame):
    frame.pack_forget()
    frame169 = Frame(root, bg=cor1)
    frame169.pack(fill="both", expand=True)
    display_logo(frame169)

    # Pergunta 2 da Chave H
    H2_label = Label(frame169, text="Especie Sabethes - Pergunta 2", font=("Times 20"), bg=cor1, fg="white")
    H2_label.pack(pady=10)
    H2_text = Text(frame169, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H2_text.insert(INSERT,
                   "A- cerda proepisternal superior presente\nB- cerda proepisternal superior ausente (subgênero Sabethes)")
    H2_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame169, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H3(frame169))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame169, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H7(frame169))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H3(frame):
    frame.pack_forget()
    frame170 = Frame(root, bg=cor1)
    frame170.pack(fill="both", expand=True)
    display_logo(frame170)

    # Pergunta 3 da Chave H
    H3_label = Label(frame170, text="Especie Sabethes - Pergunta 3", font=("Times 20"), bg=cor1, fg="white")
    H3_label.pack(pady=10)
    H3_text = Text(frame170, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H3_text.insert(INSERT,
                   "A- Tarso médio com escamas brancas (subgênero Sabethoides)\nB- Tarso médio totalmente escuro")
    H3_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame170, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H4(frame170))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame170, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H21(frame170))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H4(frame):
    frame.pack_forget()
    frame171 = Frame(root, bg=cor1)
    frame171.pack(fill="both", expand=True)
    display_logo(frame171)

    # Pergunta 4 da Chave H
    H4_label = Label(frame171, text="Especie Sabethes - Pergunta 4", font=("Times 20"), bg=cor1, fg="white")
    H4_label.pack(pady=10)
    H4_text = Text(frame171, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H4_text.insert(INSERT,
                   "A- cerda mesepímera superior curta, não alcançando a área média do mesopostnótero, alcançando apenas a área lateral\nB- cerda mesepímera superior longa, alcançando a área média do Mesoposnoto")
    H4_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame171, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H5(frame171))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame171, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H6(frame171))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H5(frame):
    frame.pack_forget()
    frame172 = Frame(root, bg=cor1)
    frame172.pack(fill="both", expand=True)
    display_logo(frame172)

    # Pergunta 5 da Chave H
    H5_label = Label(frame172, text="Especie Sabethes - Pergunta 5", font=("Times 20"), bg=cor1, fg="white")
    H5_label.pack(pady=10)
    H5_text = Text(frame172, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H5_text.insert(INSERT,
                   "A- Tarso médio com escamas brancas desde a base do tarsômero 2 até a extremidade do tarsômero 4; tergos abdominais V e VI com mancha lateral de escamas prateadas expandida em direção à extremidade como uma incisão angular\nB- Tarso médio com escamas brancas desde a base do tarsômero 3 até a extremidade ou quase até a extremidade do tarsômero 4; tergos abdominais V e VI com mancha lateral de escamas prateadas limitada à base")
    H5_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame172, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("chloropterus", frame172))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame172, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("conditus", frame172))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H6(frame):
    frame.pack_forget()
    frame173 = Frame(root, bg=cor1)
    frame173.pack(fill="both", expand=True)
    display_logo(frame173)

    # Pergunta 6 da Chave H
    H6_label = Label(frame173, text="Especie Sabethes - Pergunta 6", font=("Times 20"), bg=cor1, fg="white")
    H6_label.pack(pady=10)
    H6_text = Text(frame173, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H6_text.insert(INSERT,
                   "A- Antepronoto com escamas azuis dorsalmente\nB- Antepronoto com escamas esverdeadas dorsalmente")
    H6_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame173, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("tridentatus", frame173))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame173, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("glaucodaemon", frame173))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H7(frame):
    frame.pack_forget()
    frame174 = Frame(root, bg=cor1)
    frame174.pack(fill="both", expand=True)
    display_logo(frame174)

    # Pergunta 7 da Chave H
    H7_label = Label(frame174, text="Especie Sabethes - Pergunta 7", font=("Times 20"), bg=cor1, fg="white")
    H7_label.pack(pady=10)
    H7_text = Text(frame174, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H7_text.insert(INSERT, "A- Tarso médio totalmente escuro\nB- Tarso médio com escamas brancas")
    H7_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame174, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H8(frame174))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame174, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H15(frame174))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H8(frame):
    frame.pack_forget()
    frame175 = Frame(root, bg=cor1)
    frame175.pack(fill="both", expand=True)
    display_logo(frame175)

    # Pergunta 8 da Chave H
    H8_label = Label(frame175, text="Especie Sabethes - Pergunta 8", font=("Times 20"), bg=cor1, fg="white")
    H8_label.pack(pady=10)
    H8_text = Text(frame175, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H8_text.insert(INSERT, "A- Pernas médias sem remos\nB- Pernas médias com remos")
    H8_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame175, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("purpureus", frame175))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame175, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H9(frame175))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H9(frame):
    frame.pack_forget()
    frame176 = Frame(root, bg=cor1)
    frame176.pack(fill="both", expand=True)
    display_logo(frame176)

    # Pergunta 9 da Chave H
    H9_label = Label(frame176, text="Especie Sabethes - Pergunta 9", font=("Times 20"), bg=cor1, fg="white")
    H9_label.pack(pady=10)
    H9_text = Text(frame176, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H9_text.insert(INSERT,
                   "A- cerda mesepímera superior preta; uma única cerda amarelada pode estar presente entre as cerdas pretas\nB- cerda mesepímera superior amarelada")
    H9_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame176, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H10(frame176))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame176, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H13(frame176))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H10(frame):
    frame.pack_forget()
    frame177 = Frame(root, bg=cor1)
    frame177.pack(fill="both", expand=True)
    display_logo(frame177)

    # Pergunta 10 da Chave H
    H10_label = Label(frame177, text="Especie Sabethes - Pergunta 10", font=("Times 20"), bg=cor1, fg="white")
    H10_label.pack(pady=10)
    H10_text = Text(frame177, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H10_text.insert(INSERT,
                    "A- Pospronoto totalmente com escamas azuis\nB- Pospronoto coberto com escamas azuis dorsalmente e escamas douradas pálidas ventralmente")
    H10_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame177, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("batesi", frame177))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame177, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("cyaneus", frame177))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H11(frame):
    frame.pack_forget()
    frame178 = Frame(root, bg=cor1)
    frame178.pack(fill="both", expand=True)
    display_logo(frame178)

    # Pergunta 11 da Chave H
    H11_label = Label(frame178, text="Especie Sabethes - Pergunta 11", font=("Times 20"), bg=cor1, fg="white")
    H11_label.pack(pady=10)
    H11_text = Text(frame178, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H11_text.insert(INSERT,
                    "A- Segmentos abdominais VI e VII com cores separadas lateralmente em uma linha reta\nB- Segmentos abdominais VI e VII com cores separadas lateralmente por incisões arredondadas")
    H11_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame178, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("batesi", frame178))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame178, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("cyaneus", frame178))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H12(frame):
    frame.pack_forget()
    frame179 = Frame(root, bg=cor1)
    frame179.pack(fill="both", expand=True)
    display_logo(frame179)

    # Pergunta 12 da Chave H
    H12_label = Label(frame179, text="Especie Sabethes - Pergunta 12", font=("Times 20"), bg=cor1, fg="white")
    H12_label.pack(pady=10)
    H12_text = Text(frame179, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H12_text.insert(INSERT,
                    "A- Escutelo totalmente coberto com escamas azuis; superfície ventral da midfêmur com uma linha de escamas com reflexo dourado pálido\nB- Escutelo coberto com escamas esverdeadas, escamas das áreas antealares e supralares com reflexos dourados; superfície ventral da midfêmur escura")
    H12_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame179, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("lanei", frame179))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame179, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("spixi", frame179))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H13(frame):
    frame.pack_forget()
    frame180 = Frame(root, bg=cor1)
    frame180.pack(fill="both", expand=True)
    display_logo(frame180)

    # Pergunta 13 da Chave H
    H13_label = Label(frame180, text="Especie Sabethes - Pergunta 13", font=("Times 20"), bg=cor1, fg="white")
    H13_label.pack(pady=10)
    H13_text = Text(frame180, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H13_text.insert(INSERT,
                    "A- Antepronoto com escamas violáceas dorsalmente, contrastando com a coloração azul do escutelo\nB- Antepronoto com escamas azuis dorsalmente, semelhante à coloração do escutelo")
    H13_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame180, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("albiprivus", frame180))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame180, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H14(frame180))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H14(frame):
    frame.pack_forget()
    frame181 = Frame(root, bg=cor1)
    frame181.pack(fill="both", expand=True)
    display_logo(frame181)

    # Pergunta 14 da Chave H
    H14_label = Label(frame181, text="Especie Sabethes - Pergunta 14", font=("Times 20"), bg=cor1, fg="white")
    H14_label.pack(pady=10)
    H14_text = Text(frame181, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H14_text.insert(INSERT,
                    "A- Pospronoto totalmente com escamas azuis\nB- Pospronoto com área dorsal coberta com escamas azuis e área ventral com escamas ligeiramente douradas")
    H14_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame181, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("quasicyaneus", frame181))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame181, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("albiprivus", frame181))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H15(frame):
    frame.pack_forget()
    frame182 = Frame(root, bg=cor1)
    frame182.pack(fill="both", expand=True)
    display_logo(frame182)

    # Pergunta 15 da Chave H
    H15_label = Label(frame182, text="Especie Sabethes - Pergunta 15", font=("Times 20"), bg=cor1, fg="white")
    H15_label.pack(pady=10)
    H15_text = Text(frame182, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H15_text.insert(INSERT,
                    "A- Escutelo com faixa escura mediana distinta das áreas laterais azuis e douradas\nB- Escutelo sem faixa escura mediana, totalmente coberto com escamas da mesma cor")
    H15_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame182, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("bipartipes", frame182))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame182, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H16(frame182))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H16(frame):
    frame.pack_forget()
    frame183 = Frame(root, bg=cor1)
    frame183.pack(fill="both", expand=True)
    display_logo(frame183)

    # Pergunta 16 da Chave H
    H16_label = Label(frame183, text="Especie Sabethes - Pergunta 16", font=("Times 20"), bg=cor1, fg="white")
    H16_label.pack(pady=10)
    H16_text = Text(frame183, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H16_text.insert(INSERT,
                    "A- cerda mesepímera superior longa, alcançando a área média do Mesoposnoto\nB- cerda mesepímera superior curta, não alcançando a área média do mesopostnótero, alcançando apenas a área lateral")
    H16_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame183, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H17(frame183))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame183, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H19(frame183))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H17(frame):
    frame.pack_forget()
    frame184 = Frame(root, bg=cor1)
    frame184.pack(fill="both", expand=True)
    display_logo(frame184)

    # Pergunta 17 da Chave H
    H17_label = Label(frame184, text="Especie Sabethes - Pergunta 17", font=("Times 20"), bg=cor1, fg="white")
    H17_label.pack(pady=10)
    H17_text = Text(frame184, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H17_text.insert(INSERT,
                    "A- Mesoposnoto com mancha de escamas brancas prateadas\nB- Mesoposnoto sem mancha de escamas brancas prateadas")
    H17_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame184, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("gymnothorax", frame184))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame184, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H18(frame184))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H18(frame):
    frame.pack_forget()
    frame185 = Frame(root, bg=cor1)
    frame185.pack(fill="both", expand=True)
    display_logo(frame185)

    # Pergunta 18 da Chave H
    H18_label = Label(frame185, text="Especie Sabethes - Pergunta 18", font=("Times 20"), bg=cor1, fg="white")
    H18_label.pack(pady=10)
    H18_text = Text(frame185, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H18_text.insert(INSERT,
                    "A- Tíbia anterior com escamas brancas do tarsômero 3 até a metade basal do tarsômero 5\nB- Tíbia anterior com escamas brancas da metade apical do tarsômero 2 até o tarsômero 4")
    H18_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame185, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("amazonicus", frame185))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame185, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("forattinii", frame185))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H19(frame):
    frame.pack_forget()
    frame186 = Frame(root, bg=cor1)
    frame186.pack(fill="both", expand=True)
    display_logo(frame186)

    # Pergunta 19 da Chave H
    H19_label = Label(frame186, text="Especie Sabethes - Pergunta 19", font=("Times 20"), bg=cor1, fg="white")
    H19_label.pack(pady=10)
    H19_text = Text(frame186, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H19_text.insert(INSERT,
                    "A- Tíbia anterior totalmente escurecida no tarsômero 3\nB- Tíbia anterior com escamas brancas no lado ventral do tarsômero 3")
    H19_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame186, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("tarsopus", frame186))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame186, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H20(frame186))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H20(frame):
    frame.pack_forget()
    frame187 = Frame(root, bg=cor1)
    frame187.pack(fill="both", expand=True)
    display_logo(frame187)

    # Pergunta 20 da Chave H
    H20_label = Label(frame187, text="Especie Sabethes - Pergunta 20", font=("Times 20"), bg=cor1, fg="white")
    H20_label.pack(pady=10)
    H20_text = Text(frame187, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H20_text.insert(INSERT,
                    "A- Tíbia anterior com escamas brancas nos tarsômeros 2 e 3 no lado ventral\nB- Tíbia anterior com escamas brancas do meio apical do tarsômero 2 até a metade basal do tarsômero 4")
    H20_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame187, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("belisarioi", frame187))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame187, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("cerqueirai", frame187))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H21(frame):
    frame.pack_forget()
    frame188 = Frame(root, bg=cor1)
    frame188.pack(fill="both", expand=True)
    display_logo(frame188)

    # Pergunta 21 da Chave H
    H21_label = Label(frame188, text="Especie Sabethes - Pergunta 21", font=("Times 20"), bg=cor1, fg="white")
    H21_label.pack(pady=10)
    H21_text = Text(frame188, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H21_text.insert(INSERT,
                    "A- cerda mesocatepisternal inferior presente (subgênero Peytonulus)\nB- cerda mesocatepisternal inferior ausente (subgênero Sabethinus)")
    H21_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame188, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H22(frame188))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame188, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H30(frame188))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H22(frame):
    frame.pack_forget()
    frame189 = Frame(root, bg=cor1)
    frame189.pack(fill="both", expand=True)
    display_logo(frame189)

    # Pergunta 22 da Chave H
    H22_label = Label(frame189, text="Especie Sabethes - Pergunta 22", font=("Times 20"), bg=cor1, fg="white")
    H22_label.pack(pady=10)
    H22_text = Text(frame189, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H22_text.insert(INSERT,
                    "A- Escutelo coberto com escamas opacas, sem brilho metálico\nB- Escutelo coberto com escamas brilhantes de coloração metálica")
    H22_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame189, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("shannoni", frame189))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame189, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H23(frame189))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H23(frame):
    frame.pack_forget()
    frame190 = Frame(root, bg=cor1)
    frame190.pack(fill="both", expand=True)
    display_logo(frame190)

    # Pergunta 23 da Chave H
    H23_label = Label(frame190, text="Especie Sabethes - Pergunta 23", font=("Times 20"), bg=cor1, fg="white")
    H23_label.pack(pady=10)
    H23_text = Text(frame190, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H23_text.insert(INSERT,
                    "A- Probóscide totalmente escurecida no lado ventral\nB- Probóscide com escamas claras no lado ventral")
    H23_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame190, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H24(frame190))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame190, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H26(frame190))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H24(frame):
    frame.pack_forget()
    frame191 = Frame(root, bg=cor1)
    frame191.pack(fill="both", expand=True)
    display_logo(frame191)

    # Pergunta 24 da Chave H
    H24_label = Label(frame191, text="Especie Sabethes - Pergunta 24", font=("Times 20"), bg=cor1, fg="white")
    H24_label.pack(pady=10)
    H24_text = Text(frame191, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H24_text.insert(INSERT,
                    "A- Tarsômero 5 da tíbia posterior com escamas brancas no lado ventral\nB- Tarsômero 5 da tíbia posterior totalmente escurecido")
    H24_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame191, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H25(frame191))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame191, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("ignotus", frame191))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H25(frame):
    frame.pack_forget()
    frame192 = Frame(root, bg=cor1)
    frame192.pack(fill="both", expand=True)
    display_logo(frame192)

    # Pergunta 25 da Chave H
    H25_label = Label(frame192, text="Especie Sabethes - Pergunta 25", font=("Times 20"), bg=cor1, fg="white")
    H25_label.pack(pady=10)
    H25_text = Text(frame192, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H25_text.insert(INSERT,
                    "A- Antepronoto com escamas violáceas dorsalmente\nB- Antepronoto com escamas violáceas e azuis dorsalmente")
    H25_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame192, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("undosus", frame192))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame192, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("hadrognathus", frame192))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H26(frame):
    frame.pack_forget()
    frame193 = Frame(root, bg=cor1)
    frame193.pack(fill="both", expand=True)
    display_logo(frame193)

    # Pergunta 26 da Chave H
    H26_label = Label(frame193, text="Especie Sabethes - Pergunta 26", font=("Times 20"), bg=cor1, fg="white")
    H26_label.pack(pady=10)
    H26_text = Text(frame193, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H26_text.insert(INSERT,
                    "A- Probóscide com mancha de escamas claras no lado ventral\nB- Probóscide com linha de escamas claras no lado ventral")
    H26_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame193, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H27(frame193))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame193, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H29(frame193))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H27(frame):
    frame.pack_forget()
    frame194 = Frame(root, bg=cor1)
    frame194.pack(fill="both", expand=True)
    display_logo(frame194)

    # Pergunta 27 da Chave H
    H27_label = Label(frame194, text="Especie Sabethes - Pergunta 27", font=("Times 20"), bg=cor1, fg="white")
    H27_label.pack(pady=10)
    H27_text = Text(frame194, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H27_text.insert(INSERT,
                    "A- Tarsômero 5 da tíbia posterior com escamas brancas no lado ventral\nB- Tarsômero 5 da tíbia posterior escurecido no lado ventral")
    H27_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame194, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H28(frame194))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame194, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("whitmani", frame194))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H28(frame):
    frame.pack_forget()
    frame195 = Frame(root, bg=cor1)
    frame195.pack(fill="both", expand=True)
    display_logo(frame195)

    # Pergunta 28 da Chave H
    H28_label = Label(frame195, text="Especie Sabethes - Pergunta 28", font=("Times 20"), bg=cor1, fg="white")
    H28_label.pack(pady=10)
    H28_text = Text(frame195, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H28_text.insert(INSERT,
                    "A- Antepronoto com escamas violáceas dorsalmente\nB- Antepronoto com escamas azuis dorsalmente")
    H28_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame195, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("soperi", frame195))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame195, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("fabricii", frame195))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H29(frame):
    frame.pack_forget()
    frame196 = Frame(root, bg=cor1)
    frame196.pack(fill="both", expand=True)
    display_logo(frame196)

    # Pergunta 29 da Chave H
    H29_label = Label(frame196, text="Especie Sabethes - Pergunta 29", font=("Times 20"), bg=cor1, fg="white")
    H29_label.pack(pady=10)
    H29_text = Text(frame196, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H29_text.insert(INSERT,
                    "A- Tarsômero 5 da tíbia posterior com escamas brancas no lado ventral\nB- Tarsômero 5 da tíbia posterior totalmente escurecido")
    H29_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame196, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("aurescens", frame196))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame196, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("identicus", frame196))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H30(frame):
    frame.pack_forget()
    frame197 = Frame(root, bg=cor1)
    frame197.pack(fill="both", expand=True)
    display_logo(frame197)

    # Pergunta 30 da Chave H
    H30_label = Label(frame197, text="Especie Sabethes - Pergunta 30", font=("Times 20"), bg=cor1, fg="white")
    H30_label.pack(pady=10)
    H30_text = Text(frame197, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H30_text.insert(INSERT,
                    "A- Vértebra e área dorsal do antepronoto com a mesma coloração metálica\nB- Vértebra e área dorsal do antepronoto com coloração metálica distinta")
    H30_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame197, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H31(frame197))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame197, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_H32(frame197))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H31(frame):
    frame.pack_forget()
    frame198 = Frame(root, bg=cor1)
    frame198.pack(fill="both", expand=True)
    display_logo(frame198)

    # Pergunta 31 da Chave H
    H31_label = Label(frame198, text="Especie Sabethes - Pergunta 31", font=("Times 20"), bg=cor1, fg="white")
    H31_label.pack(pady=10)
    H31_text = Text(frame198, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H31_text.insert(INSERT,
                    "A- Vértebra e área dorsal do antepronoto com reflexos violáceos\nB- Vértebra e área dorsal do antepronoto com reflexos azulados")
    H31_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame198, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("intermedius", frame198))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame198, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("melanonymphe", frame198))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_H32(frame):
    frame.pack_forget()
    frame199 = Frame(root, bg=cor1)
    frame199.pack(fill="both", expand=True)
    display_logo(frame199)

    # Pergunta 32 da Chave H
    H32_label = Label(frame199, text="Especie Sabethes - Pergunta 32", font=("Times 20"), bg=cor1, fg="white")
    H32_label.pack(pady=10)
    H32_text = Text(frame199, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    H32_text.insert(INSERT,
                    "A- Vértebra predominantemente com reflexos violáceos e área dorsal do antepronoto com reflexos azul-claro\nB- Vértebra predominantemente com reflexos violáceos e antepronoto com alguns reflexos violáceos")
    H32_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame199, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("xhyphydes", frame199))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame199, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_especie("idiogene", frame199))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def create_chave_i_widgets():
    frame200 = Frame(root, bg=cor1)
    frame200.pack(fill="both", expand=True)
    display_logo(frame200)

    # Pergunta 1 da Chave H
    I1_label = Label(frame200, text="Gêneros Culicini - Pergunta 1", font=("Times 20"), bg=cor1, fg="white")
    I1_label.pack(pady=10)
    I1_text = Text(frame200, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I1_text.insert(INSERT,
                   "A - Cerdas mesepimerais inferiores (SMnti) presentes, formando fileira com cerca de cinco elementos desenvolvidos; fêmures salpicados de escamas claras; costa (C) com áreas claras e escuras, alternadas; mosquitos de porte avantajado  \n \nB-Cerdas mesepimerais inferiores (SMmi) ausentes ou, se presentes, finas e em número reduzido, às vezes limitadas a um único elemento; fêmures não conspicuamente salpicados de escamas claras; costa (C) sem áreas claras e escuras alternadas; mosquitos de porte médio ou pequeno  ")
    I1_text.pack(pady=10)

    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame200, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Lutzia ", frame200))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame200, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I2(frame200))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I2(frame):
    frame.pack_forget()
    frame201 = Frame(root, bg=cor1)
    frame201.pack(fill="both", expand=True)
    display_logo(frame201)

    # Pergunta 2 da Chave I
    I2_label = Label(frame201, text="Genero Culicini - Pergunta 2", font=("Times 20"), bg=cor1, fg="white")
    I2_label.pack(pady=10)
    I2_text = Text(frame201, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I2_text.insert(INSERT,
                   "A- Abdome com os tergitos dotados de áreas basolaterais de escamas claras e com reflexos metálicos\n \nB- Abdome com tergitos escuros ou com áreas basolaterais, ou faixas basais, de escamas claras mas sem reflexos metálicos")
    I2_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame201, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Carrollia", frame201))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame201, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I3(frame201))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I3(frame):
    frame.pack_forget()
    frame202 = Frame(root, bg=cor1)
    frame202.pack(fill="both", expand=True)
    display_logo(frame202)

    # Pergunta 3 da Chave I
    I3_label = Label(frame202, text="Genero Culicini - Pergunta 3", font=("Times 20"), bg=cor1, fg="white")
    I3_label.pack(pady=10)
    I3_text = Text(frame202, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I3_text.insert(INSERT,
                   "A- Cerdas acrosticais (SAc) desenvolvidas e situadas na área central do escudo (Esc)\n \nB- Cerdas acrosticais (SAc) ausentes na área central do escudo (Esc) e, se presentes, confinadas à extremidade anterior da área ou, então, também na região pré-escutelar (Apr)")
    I3_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame202, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I4(frame202))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame202, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I9(frame202))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I4(frame):
    frame.pack_forget()
    frame203 = Frame(root, bg=cor1)
    frame203.pack(fill="both", expand=True)
    display_logo(frame203)

    # Pergunta 4 da Chave I
    I4_label = Label(frame203, text="Genero Culicini - Pergunta 4", font=("Times 20"), bg=cor1, fg="white")
    I4_label.pack(pady=10)
    I4_text = Text(frame203, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I4_text.insert(INSERT,
                   "A- Mesocatepisterno (Mcs) com escamas em manchas distintas\n \nB- Mesocatepisterno (Mcs) com apenas algumas escamas entre as cerdas")
    I4_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame203, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I5(frame203))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame203, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I7(frame203))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I5(frame):
    frame.pack_forget()
    frame204 = Frame(root, bg=cor1)
    frame204.pack(fill="both", expand=True)
    display_logo(frame204)

    # Pergunta 5 da Chave I
    I5_label = Label(frame204, text="Genero Culicini - Pergunta 5", font=("Times 20"), bg=cor1, fg="white")
    I5_label.pack(pady=10)
    I5_text = Text(frame204, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I5_text.insert(INSERT,
                   "A- Tarsômero posterior Ta-1 mais curto do que a tíbia correspondente\n \nB- Tarsômero posterior Ta-1 de comprimento subigual ou mais longo do que a tíbia correspondente")
    I5_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame204, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Allimanta", frame204))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame204, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I6(frame204))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I6(frame):
    frame.pack_forget()
    frame205 = Frame(root, bg=cor1)
    frame205.pack(fill="both", expand=True)
    display_logo(frame205)

    # Pergunta 6 da Chave I
    I6_label = Label(frame205, text="Genero Culicini - Pergunta 6", font=("Times 20"), bg=cor1, fg="white")
    I6_label.pack(pady=10)
    I6_text = Text(frame205, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I6_text.insert(INSERT,
                   "A- Escudo (Esc) com faixa de escamas claras, estendida lateralmente, da área acrostical (AA) até a supra-alar (ASa)\n \nB- Escudo (Esc) coberto de escamas geralmente unicoloridas e, quando ocorrem faixas ou áreas claras, são pouco evidentes")
    I6_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame205, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Phenacomyia", frame205))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame205, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Culex", frame205))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I7(frame):
    frame.pack_forget()
    frame206 = Frame(root, bg=cor1)
    frame206.pack(fill="both", expand=True)
    display_logo(frame206)

    # Pergunta 7 da Chave I
    I7_label = Label(frame206, text="Genero Culicini - Pergunta 7", font=("Times 20"), bg=cor1, fg="white")
    I7_label.pack(pady=10)
    I7_text = Text(frame206, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I7_text.insert(INSERT,
                   "A- Pleura com tegumento amarelado ou esbranquiçado\n \nB- Pleura com tegumento escuro ou claro mas, neste último caso, com áreas escurecidas")
    I7_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame206, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Micraedes", frame206))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame206, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I8(frame206))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I8(frame):
    frame.pack_forget()
    frame207 = Frame(root, bg=cor1)
    frame207.pack(fill="both", expand=True)
    display_logo(frame207)

    # Pergunta 8 da Chave I
    I8_label = Label(frame207, text="Genero Culicini - Pergunta 8", font=("Times 20"), bg=cor1, fg="white")
    I8_label.pack(pady=10)
    I8_text = Text(frame207, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I8_text.insert(INSERT, "A- Tarsômeros com área clara basal\n \nB- Tarsômeros completamente escuros")
    I8_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame207, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Microculex", frame207))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame207, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Anoedioporpa", frame207))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I9(frame):
    frame.pack_forget()
    frame208 = Frame(root, bg=cor1)
    frame208.pack(fill="both", expand=True)
    display_logo(frame208)

    # Pergunta 9 da Chave I
    I9_label = Label(frame208, text="Genero Culicini - Pergunta 9", font=("Times 20"), bg=cor1, fg="white")
    I9_label.pack(pady=10)
    I9_text = Text(frame208, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I9_text.insert(INSERT,
                   "A- Abdome com os tergitos apresentando áreas claras apicais ou apicolaterais\n \nB- Abdome com os tergitos totalmente escuros ou com áreas claras basais ou basolaterais")
    I9_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame208, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Neoculex", frame208))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame208, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I10(frame208))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I10(frame):
    frame.pack_forget()
    frame209 = Frame(root, bg=cor1)
    frame209.pack(fill="both", expand=True)
    display_logo(frame209)

    # Pergunta 10 da Chave I
    I10_label = Label(frame209, text="Genero Culicini - Pergunta 10", font=("Times 20"), bg=cor1, fg="white")
    I10_label.pack(pady=10)
    I10_text = Text(frame209, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I10_text.insert(INSERT,
                    "A- Vértice (V) com escamas largas ao longo da linha limítrofe entre os olhos\n \nB- Vértice (V) com escamas estreitas ao longo da linha limítrofe entre os olhos")
    I10_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame209, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I11(frame209))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame209, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I12(frame209))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I11(frame):
    frame.pack_forget()
    frame210 = Frame(root, bg=cor1)
    frame210.pack(fill="both", expand=True)
    display_logo(frame210)

    # Pergunta 11 da Chave I
    I11_label = Label(frame210, text="Genero Culicini - Pergunta 11", font=("Times 20"), bg=cor1, fg="white")
    I11_label.pack(pady=10)
    I11_text = Text(frame210, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I11_text.insert(INSERT,
                    "A- Indivíduos masculinos com os palpos maxilares (PlpM) longos, mais compridos do que a probóscide\n \nB- Indivíduos masculinos com os palpos maxilares (PlpM) curtos, de comprimento inferior ao da probóscide")
    I11_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame210, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Melanoconion", frame210))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame210, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Aedinus", frame210))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I12(frame):
    frame.pack_forget()
    frame211 = Frame(root, bg=cor1)
    frame211.pack(fill="both", expand=True)
    display_logo(frame211)

    # Pergunta 12 da Chave I
    I12_label = Label(frame211, text="Genero Culicini - Pergunta 12", font=("Times 20"), bg=cor1, fg="white")
    I12_label.pack(pady=10)
    I12_text = Text(frame211, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I12_text.insert(INSERT, "A- Mesanepímero (Man) sem cerdas\n \nB- Mesanepímero (Mani) com cerdas, geralmente poucas")
    I12_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame211, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Aedinus", frame211))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame211, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I13(frame211))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I13(frame):
    frame.pack_forget()
    frame212 = Frame(root, bg=cor1)
    frame212.pack(fill="both", expand=True)
    display_logo(frame212)

    # Pergunta 13 da Chave I
    I13_label = Label(frame212, text="Genero Culicini - Pergunta 13", font=("Times 20"), bg=cor1, fg="white")
    I13_label.pack(pady=10)
    I13_text = Text(frame212, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I13_text.insert(INSERT,
                    "A- Cerdas proepisternais superiores (SPeS) desenvolvidas e em número apreciável\n \nB- Cerdas proepisternais superiores (SPeS) pouco desenvolvidas e em pequeno número")
    I13_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame212, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Tinolestes", frame212))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame212, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I14(frame212))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)


def show_question_I14(frame):
    frame.pack_forget()
    frame213 = Frame(root, bg=cor1)
    frame213.pack(fill="both", expand=True)
    display_logo(frame213)

    # Pergunta 14 da Chave I
    I14_label = Label(frame213, text="Genero Culicini - Pergunta 14", font=("Times 20"), bg=cor1, fg="white")
    I14_label.pack(pady=10)
    I14_text = Text(frame213, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I14_text.insert(INSERT, "A- Veias R2 e R3 com escamas largas\n \nB- Veias R2 e R3 com escamas estreitas")
    I14_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame213, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Aedinus", frame213))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame213, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: show_question_I15(frame213))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)

def show_question_I15(frame):
    frame.pack_forget()
    frame214 = Frame(root, bg=cor1)
    frame214.pack(fill="both", expand=True)
    display_logo(frame214)

    # Pergunta 15 da Chave I
    I15_label = Label(frame214, text="Genero Culicini - Pergunta 15", font=("Times 20"), bg=cor1, fg="white")
    I15_label.pack(pady=10)
    I15_text = Text(frame214, height=10, width=80, font=("Times 15"), bg=cor1, fg="white", wrap='word')
    I15_text.insert(INSERT,
                    "A- Mesocatepisterno (Mcs) com distinta mancha de escamas\n \nB- Mesocatepisterno (Mcs) sem distinta mancha de escamas")
    I15_text.pack(pady=10)

    # Carregando e redimensionando as imagens
    image_A = Image.open("ArboTeamPng.png").resize((400, 400))
    image_B = Image.open("ArboTeamPng.png").resize((400, 400))

    image_A = ImageTk.PhotoImage(image_A)
    image_B = ImageTk.PhotoImage(image_B)

    button_A = Button(frame214, text="A", image=image_A, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Microculex", frame214))
    button_A.image = image_A
    button_A.pack(side="left", padx=10)

    button_B = Button(frame214, text="B", image=image_B, compound="left", font=("Times 15"), bg=cor2, fg="white",
                      command=lambda: set_genero("Belkiriomyia", frame214))
    button_B.image = image_B
    button_B.pack(side="right", padx=10)





create_save_method_widgets()
root.mainloop()
