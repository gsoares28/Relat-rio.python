import tkinter as tk
from tkinter import ttk
import pandas as pd

# Nome do arquivo Excel
arquivo_excel = "despesas.xlsx"

# Carregar os dados do Excel para um DataFrame
try:
    df = pd.read_excel(arquivo_excel)
except FileNotFoundError:
    # Cria um DataFrame vazio se o arquivo não existir
    df = pd.DataFrame(columns=["Categoria", "Vencimento", "Status"])

# Função para atualizar a planilha após edições
def atualizar_planilha():
    # Extrair dados do Treeview e atualizar o DataFrame
    dados_atualizados = []
    for row in treeview.get_children():
        valores = treeview.item(row, "values")
        dados_atualizados.append(valores)
    
    # Atualizar o DataFrame com os dados extraídos
    colunas = ["Categoria", "Vencimento", "Status"]
    df_atualizado = pd.DataFrame(dados_atualizados, columns=colunas)
    
    # Salvar no arquivo Excel
    df_atualizado.to_excel(arquivo_excel, index=False)
    print("Planilha salva com sucesso!")

# Função para permitir a edição de uma linha no Treeview
def on_select(event):
    selected_item = treeview.selection()[0]
    valores = treeview.item(selected_item, "values")
    
    # Preencher os campos de edição
    categoria_var.set(valores[0])
    vencimento_var.set(valores[1])
    status_var.set(valores[2])

# Função para salvar alterações feitas na linha selecionada
def salvar_alteracoes():
    selected_item = treeview.selection()[0]
    
    # Obter valores editados
    categoria = categoria_var.get()
    vencimento = vencimento_var.get()
    status = status_var.get()
    
    # Atualizar a linha selecionada no Treeview
    treeview.item(selected_item, values=(categoria, vencimento, status))

# Função chamada ao fechar o aplicativo
def ao_fechar():
    atualizar_planilha()  # Salva as alterações no arquivo Excel
    root.destroy()  # Fecha a janela

# Configuração da janela principal
root = tk.Tk()
root.title("Controle de Despesas")

# Criar o Treeview (tabela)
treeview = ttk.Treeview(root, columns=("Categoria", "Vencimento", "Status"), show="headings")

# Configurar as colunas do Treeview
treeview.heading("Categoria", text="Categoria")
treeview.heading("Vencimento", text="Vencimento")
treeview.heading("Status", text="Status")

# Preencher o Treeview com os dados do DataFrame
for _, row in df.iterrows():
    treeview.insert("", "end", values=(row["Categoria"], row["Vencimento"], row["Status"]))

treeview.pack(padx=20, pady=20)

# Variáveis para os campos de edição
categoria_var = tk.StringVar()
vencimento_var = tk.StringVar()
status_var = tk.StringVar()

# Criar campos de edição
frame_edicao = tk.Frame(root)
frame_edicao.pack(pady=20)

tk.Label(frame_edicao, text="Categoria").grid(row=0, column=0)
tk.Entry(frame_edicao, textvariable=categoria_var).grid(row=0, column=1)

tk.Label(frame_edicao, text="Vencimento").grid(row=1, column=0)
tk.Entry(frame_edicao, textvariable=vencimento_var).grid(row=1, column=1)

tk.Label(frame_edicao, text="Status").grid(row=2, column=0)
tk.Entry(frame_edicao, textvariable=status_var).grid(row=2, column=1)

# Botão para salvar as alterações
btn_salvar = tk.Button(root, text="Salvar Alterações", command=salvar_alteracoes)
btn_salvar.pack(pady=10)

# Evento de seleção na tabela
treeview.bind("<<TreeviewSelect>>", on_select)

# Configurar o evento de fechamento da janela
root.protocol("WM_DELETE_WINDOW", ao_fechar)

# Iniciar a interface gráfica
root.mainloop()
