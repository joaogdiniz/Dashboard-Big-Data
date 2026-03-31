import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import threading
from queue import Queue
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.ticker as ticker

# Configuração Inicial da Janela
root = tk.Tk()
root.withdraw()
root.title("Membros Autistas Botafoguenses - Dashboard")

# Tenta iniciar maximizado ou com tamanho seguro
try:
    root.state('zoomed')
except:
    root.geometry("1280x720")

CSV_FILE = "dados.csv"
# URL de Exportação
URL_SHEET = "https://docs.google.com/spreadsheets/d/1eba-a2wLFpPUkdlfD9teaOLth8Mlx1LZI8nBGr4t3Rc/export?format=csv&gid=779899065"

COLUNAS_PRINCIPAIS = [
    "ID", "Nome", "Nascimento", "Telefone", "CPF", "Email", "Profissão",
    "Carteirinha", "Ajuda", "PCD", "TEA", "CEP", "DataEntrada"
]

sync_queue = Queue()
INTERVALO_SYNC_MS = 300000  # 5 minutos

# Variáveis globais para janelas de gráficos
janela_graf_cart = None
janela_graf_tea = None
janela_graf_crescimento = None
janela_graf_ajuda_linha = None

def sincronizar_dados(mostrar_popup=True):
    df_local = pd.DataFrame(columns=COLUNAS_PRINCIPAIS)
    
    # 1. Carregar Local
    if os.path.exists(CSV_FILE):
        try:
            df_local = pd.read_csv(CSV_FILE)
            if "Email" in df_local.columns:
                df_local["Email"] = df_local["Email"].astype(str).str.lower().str.strip()
        except Exception as e:
            print(f"Erro ao ler CSV local: {e}")

    # 2. Carregar Remoto
    try:
        df_remoto = pd.read_csv(URL_SHEET)
        # Limpeza de nomes de colunas
        df_remoto.columns = df_remoto.columns.str.replace(r'\s+', ' ', regex=True).str.strip()
        
        mapa_colunas = {
            "Nome completo": "Nome",
            "Endereço de e-mail": "Email",
            "Deseja receber a carteirinha ?": "Carteirinha",
            "É PCD - se sim, descreva": "PCD",
            "Telefone para contato Ex: 21 9 9999-9999": "Telefone",
            "Trabalha com o público TEA ?": "TEA",
            "Data de nascimento": "Nascimento",
            "Deseja nos ajudar pagando apenas 15 R$ mensais e obter descontos exlusivos ?": "Ajuda",
            "CPF ( Obrigatório somente para membros que escolherem pagar mensalmente )": "CPF",
            "Carimbo de data/hora": "DataEntrada",
            "Endereço: CEP": "CEP"
        }
        df_remoto.rename(columns=mapa_colunas, inplace=True)
        
        # Padronização
        if "DataEntrada" in df_remoto.columns:
            df_remoto["DataEntrada"] = pd.to_datetime(df_remoto["DataEntrada"], errors="coerce")
        
        if "Email" in df_remoto.columns:
            df_remoto["Email"] = df_remoto["Email"].astype(str).str.lower().str.strip()

        if "Carteirinha" in df_remoto.columns:
            df_remoto["Carteirinha"] = df_remoto["Carteirinha"].astype(str).str[:3]

    except Exception as e:
        print(f"AVISO: Não foi possível carregar dados da planilha. Erro: {e}")
        if mostrar_popup:
            messagebox.showwarning("Modo Offline", "Sem conexão com a planilha Google. Usando dados locais.")
        return df_local

    # 3. Combinar e Remover Duplicatas
    # Prioriza dados do df_remoto se houver conflito
    df_combinado = pd.concat([df_remoto, df_local], ignore_index=True)
    
    # Remove duplicatas baseado no Email
    if "Email" in df_combinado.columns:
        df_combinado = df_combinado.drop_duplicates(subset=["Email"], keep="first")
    
    # Filtra apenas colunas desejadas que existam no dataframe
    cols_existentes = [col for col in COLUNAS_PRINCIPAIS if col in df_combinado.columns]
    df_combinado = df_combinado[cols_existentes]

    try:
        df_combinado.to_csv(CSV_FILE, index=False)
        if mostrar_popup:
            messagebox.showinfo("Sincronização", f"Dados atualizados!\nTotal: {len(df_combinado)} registros.")
    except Exception as e:
        print(f"Erro ao salvar CSV: {e}")

    # Regenera IDs sequenciais para visualização
    df_combinado = df_combinado.reset_index(drop=True)
    df_combinado["ID"] = df_combinado.index + 1
    
    return df_combinado

def sincronizar_em_background():
    print("Iniciando sync background...")
    try:
        df_novo = sincronizar_dados(mostrar_popup=False)
        if df_novo is not None and not df_novo.empty:
            sync_queue.put(df_novo)
            print("Sync background concluído.")
    except Exception as e:
        print(f"Erro no background: {e}")

def agendar_sincronizacao_periodica():
    threading.Thread(target=sincronizar_em_background, daemon=True).start()
    root.after(INTERVALO_SYNC_MS, agendar_sincronizacao_periodica)

def verificar_fila_e_atualizar_ui():
    global df
    try:
        df_novo = sync_queue.get(block=False)
        # Verifica se mudou o tamanho ou algo relevante para não redesenhar à toa
        if len(df_novo) != len(df) or not df_novo.equals(df):
            print("Atualizando UI com novos dados...")
            df = df_novo
            atualizar_tabela(df)
    except Exception:
        pass
    root.after(1000, verificar_fila_e_atualizar_ui) # Verifica a cada 1s

# --- Inicialização dos Dados ---
df = sincronizar_dados(mostrar_popup=True)

# --- Lógica de CRUD Local ---
def salvar_registro():
    global df
    novo = {
        "Nome": entry_nome.get(),
        "Nascimento": entry_nascimento.get(),
        "Email": entry_email.get().lower().strip(),
        "Telefone": entry_telefone.get(),
        "CPF": entry_cpf.get(),
        "Profissão": entry_profissao.get(),
        "Carteirinha": entry_carteirinha.get(),
        "Ajuda": entry_ajuda.get(),
        "PCD": entry_pcd.get(),
        "TEA": entry_tea.get(),
        "CEP": entry_cep.get(),
        "DataEntrada": pd.Timestamp.now()
    }
    
    if not novo["Email"]:
        messagebox.showwarning("Erro", "O Email é obrigatório.")
        return

    if not df.empty and novo["Email"] in df["Email"].values:
        messagebox.showwarning("Erro", "Este email já está cadastrado.")
    else:
        novo_df = pd.DataFrame([novo])
        df = pd.concat([df, novo_df], ignore_index=True)
        df["ID"] = df.index + 1
        df.to_csv(CSV_FILE, index=False)
        messagebox.showinfo("Sucesso", "Registro adicionado!")
        atualizar_tabela(df)
        # Limpar campos
        for entry in entries: entry.delete(0, tk.END)

def atualizar_tabela(dados):
    for i in tabela.get_children():
        tabela.delete(i)
    
    if dados is None or dados.empty:
        return

    dados_display = dados.copy()
    # Garante que colunas existam
    for col in COLUNAS_PRINCIPAIS:
        if col not in dados_display.columns:
            dados_display[col] = ""
            
    # Formata data para string
    if "DataEntrada" in dados_display.columns:
        dados_display["DataEntrada"] = pd.to_datetime(dados_display["DataEntrada"], errors='coerce').dt.strftime('%Y-%m-%d')

    for idx, row in dados_display.iterrows():
        values = [row.get(col, "") for col in COLUNAS_PRINCIPAIS]
        tabela.insert("", "end", iid=idx, values=values)

def excluir():
    selected_items = tabela.selection()
    if not selected_items:
        messagebox.showwarning("Atenção", "Selecione um registro para excluir")
        return
    global df
    if messagebox.askyesno("Confirmar", "Deseja realmente excluir o(s) registro(s)?"):
        try:
            indices_tabela = [int(item) for item in selected_items]
            
            emails_para_remover = []
            for i in selected_items:
                vals = tabela.item(i)['values']
                emails_para_remover.append(str(vals[5])) 

            df = df[~df["Email"].isin(emails_para_remover)]
            df.to_csv(CSV_FILE, index=False)
            atualizar_tabela(df)
            messagebox.showinfo("Sucesso", "Registro(s) excluído(s).")
        except Exception as e:
            print(f"Erro ao excluir: {e}")
            messagebox.showerror("Erro", "Erro ao excluir.")

def pesquisar():
    coluna = coluna_var.get()
    valor = entrada_valor.get()
    if coluna and valor and not df.empty:
        try:
            resultado = df[df[coluna].astype(str).str.contains(valor, case=False, na=False)]
            atualizar_tabela(resultado)
        except Exception as e:
            print(f"Erro na busca: {e}")
    else:
        atualizar_tabela(df)

# --- Interface Gráfica ---
frame_cadastro = tk.LabelFrame(root, text="Cadastro de Membro")
frame_cadastro.pack(fill="x", padx=10, pady=5)

labels = ["Nome", "Nascimento", "Email", "Telefone", "CPF","Profissão","Carteirinha","Ajuda","PCD","TEA","CEP"]
entries_widgets = {}
entries = []

# Quebra linha a cada 4 itens
cols_per_row = 6
for i, label_text in enumerate(labels):
    r = (i // cols_per_row) * 2
    c = i % cols_per_row
    
    ttk.Label(frame_cadastro, text=label_text).grid(row=r, column=c, padx=5, pady=(5,0), sticky="w")
    entry = ttk.Entry(frame_cadastro)
    entry.grid(row=r+1, column=c, padx=5, pady=(0,5), sticky="ew")
    entries.append(entry)
    
    # Referencias globais para salvar_registro usar
    if label_text == "Nome": entry_nome = entry
    elif label_text == "Nascimento": entry_nascimento = entry
    elif label_text == "Email": entry_email = entry
    elif label_text == "Telefone": entry_telefone = entry
    elif label_text == "CPF": entry_cpf = entry
    elif label_text == "Profissão": entry_profissao = entry
    elif label_text == "Carteirinha": entry_carteirinha = entry
    elif label_text == "Ajuda": entry_ajuda = entry
    elif label_text == "PCD": entry_pcd = entry
    elif label_text == "TEA": entry_tea = entry
    elif label_text == "CEP": entry_cep = entry

ttk.Button(frame_cadastro, text="Salvar no Banco Local", command=salvar_registro).grid(row=4, column=0, columnspan=2, pady=10, sticky="w", padx=5)

# --- Área de Controles e Gráficos ---
frame_controles = tk.Frame(root)
frame_controles.pack(fill="x", padx=10)

frame_pesquisa = tk.LabelFrame(frame_controles, text="Pesquisa")
frame_pesquisa.pack(side="left", fill="y", padx=(0, 10))

coluna_var = tk.StringVar()
coluna_menu = ttk.Combobox(frame_pesquisa, textvariable=coluna_var, values=COLUNAS_PRINCIPAIS, state="readonly", width=15)
if COLUNAS_PRINCIPAIS: coluna_menu.current(1) # Seleciona Nome por padrão
coluna_menu.pack(side="left", padx=5, pady=5)

entrada_valor = ttk.Entry(frame_pesquisa, width=20)
entrada_valor.pack(side="left", padx=5, pady=5)
ttk.Button(frame_pesquisa, text="Buscar", command=pesquisar).pack(side="left", padx=5)
ttk.Button(frame_pesquisa, text="Limpar", command=lambda: atualizar_tabela(df)).pack(side="left", padx=5)

frame_graficos = tk.LabelFrame(frame_controles, text="Dashboard Analytics")
frame_graficos.pack(side="left", fill="both", expand=True)

# --- Funções dos Gráficos (Mantidas e Corrigidas) ---
def abrir_grafico_carteirinha():
    global janela_graf_cart
    if "Carteirinha" not in df.columns or df.empty: return
    
    if janela_graf_cart and janela_graf_cart.winfo_exists():
        janela_graf_cart.lift()
        return

    janela_graf_cart = tk.Toplevel(root)
    janela_graf_cart.title("Solicitações de Carteirinha")
    janela_graf_cart.geometry("600x400")
    
    # Normalizar dados (Sim/Não/Talvez)
    dados = df["Carteirinha"].str.lower().str.strip()
    contagem = dados.value_counts()
    
    fig = Figure(figsize=(6, 4), dpi=100)
    ax = fig.add_subplot(111)
    barras = ax.bar(contagem.index, contagem.values, color='#4CAF50')
    ax.set_title("Interesse na Carteirinha")
    ax.bar_label(barras)
    
    FigureCanvasTkAgg(fig, master=janela_graf_cart).get_tk_widget().pack(fill="both", expand=True)

def abrir_grafico_tea():
    global janela_graf_tea
    if "TEA" not in df.columns or df.empty: return

    if janela_graf_tea and janela_graf_tea.winfo_exists():
        janela_graf_tea.lift()
        return

    janela_graf_tea = tk.Toplevel(root)
    janela_graf_tea.title("Profissionais TEA")
    
    contagem = df["TEA"].value_counts()
    fig = Figure(figsize=(5, 5), dpi=100)
    ax = fig.add_subplot(111)
    ax.pie(contagem, labels=contagem.index, autopct='%1.1f%%', startangle=90)
    ax.set_title("Trabalha com público TEA?")
    
    FigureCanvasTkAgg(fig, master=janela_graf_tea).get_tk_widget().pack(fill="both", expand=True)

def abrir_grafico_crescimento():
    global janela_graf_crescimento
    if janela_graf_crescimento and janela_graf_crescimento.winfo_exists():
        janela_graf_crescimento.lift()
        return

    # Preparação de dados segura
    temp_df = df.copy()
    if "DataEntrada" not in temp_df.columns: return
    
    temp_df["DataEntrada"] = pd.to_datetime(temp_df["DataEntrada"], errors="coerce")
    temp_df = temp_df.dropna(subset=["DataEntrada"])
    if temp_df.empty: 
        messagebox.showinfo("Info", "Sem datas válidas para gerar gráfico de crescimento.")
        return

    temp_df["Ano"] = temp_df["DataEntrada"].dt.year
    crescimento = temp_df.groupby("Ano").size().cumsum().reset_index(name="Total")
    
    janela_graf_crescimento = tk.Toplevel(root)
    janela_graf_crescimento.title("Crescimento Acumulado")
    
    fig = Figure(figsize=(6, 4), dpi=100)
    ax = fig.add_subplot(111)
    ax.plot(crescimento["Ano"], crescimento["Total"], marker='o', linestyle='-', color='blue')
    ax.set_title("Evolução do Número de Membros")
    ax.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    ax.grid(True, linestyle='--')
    
    FigureCanvasTkAgg(fig, master=janela_graf_crescimento).get_tk_widget().pack(fill="both", expand=True)

def abrir_grafico_ajuda_linha():
    global janela_graf_ajuda_linha
    if "Ajuda" not in df.columns or "DataEntrada" not in df.columns: return
    
    if janela_graf_ajuda_linha and janela_graf_ajuda_linha.winfo_exists():
        janela_graf_ajuda_linha.lift()
        return

    # Processamento de dados
    temp_df = df.copy()
    temp_df["DataEntrada"] = pd.to_datetime(temp_df["DataEntrada"], errors="coerce")
    temp_df = temp_df.dropna(subset=["DataEntrada"])
    temp_df["Ano"] = temp_df["DataEntrada"].dt.year
    
    # Filtra Sim/Não
    temp_df["Ajuda_Sim"] = temp_df["Ajuda"].astype(str).str.lower().str.contains("sim")
    
    agrupado = temp_df.groupby("Ano")["Ajuda_Sim"].sum().reset_index(name="Doadores")
    
    janela_graf_ajuda_linha = tk.Toplevel(root)
    janela_graf_ajuda_linha.title("Evolução de Doadores")
    
    fig = Figure(figsize=(6, 4), dpi=100)
    ax = fig.add_subplot(111)
    ax.plot(agrupado["Ano"], agrupado["Doadores"], marker='s', color='green', label="Membros Contribuintes")
    ax.set_title("Novos Doadores por Ano")
    ax.legend()
    ax.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    
    FigureCanvasTkAgg(fig, master=janela_graf_ajuda_linha).get_tk_widget().pack(fill="both", expand=True)

# Botões de Gráfico
ttk.Button(frame_graficos, text="Carteirinha (Barras)", command=abrir_grafico_carteirinha).pack(side="left", padx=5, pady=5)
ttk.Button(frame_graficos, text="TEA (Pizza)", command=abrir_grafico_tea).pack(side="left", padx=5, pady=5)
ttk.Button(frame_graficos, text="Crescimento (Linha)", command=abrir_grafico_crescimento).pack(side="left", padx=5, pady=5)
ttk.Button(frame_graficos, text="Doações (Anual)", command=abrir_grafico_ajuda_linha).pack(side="left", padx=5, pady=5)

# --- Tabela ---
frame_tabela = tk.Frame(root)
frame_tabela.pack(fill="both", expand=True, padx=10, pady=5)

tabela_scroll = ttk.Scrollbar(frame_tabela)
tabela_scroll.pack(side="right", fill="y")

tabela = ttk.Treeview(frame_tabela, columns=COLUNAS_PRINCIPAIS, show="headings", yscrollcommand=tabela_scroll.set)
tabela_scroll.config(command=tabela.yview)

for col in COLUNAS_PRINCIPAIS:
    tabela.heading(col, text=col)
    tabela.column(col, width=100, minwidth=50)

tabela.pack(fill="both", expand=True)
ttk.Button(root, text="Excluir Registro Selecionado", command=excluir).pack(pady=5)

# --- Finalização ---
def ao_fechar():
    # NÃO APAGAR O CSV - Isso garante que funcione offline na próxima vez
    root.destroy()

root.protocol("WM_DELETE_WINDOW", ao_fechar)
root.deiconify()
verificar_fila_e_atualizar_ui()
agendar_sincronizacao_periodica()

root.mainloop()