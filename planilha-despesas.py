import pandas as pd

# Definir as categorias e vencimentos
categorias = [
    "Carro", "Aluguel Casa", "Aluguel Salão", "Luz Casa", "Luz Salão",
    "Internet Casa", "Internet Salão", "Celular", "Compras Casa", "Plano de Saúde",
    "Faculdade", "Compras Produtos Salão", "Cartão Nubank Salão", "Cartão Gus",
    "Cartão Nubank Carla", "Cartão Salão PagSeguro", "Cartão PagSeguro Gus",
    "Academia", "Produtos de Academia", "Lanches FDS", "Gasolina", 
    "Dinheiro para Emergência", "Lavar Carro"
]

vencimentos = [
    "15", "10", "05", "25", "25", 
    "10", "10", "05", "30", "05", 
    "20", "15", "10", "15", "10", 
    "20", "15", "25", "30", "28", 
    "05", "20","05"
]

# Verificar o número de elementos em cada lista
print(f"Tamanho da lista de categorias: {len(categorias)}")
print(f"Tamanho da lista de vencimentos: {len(vencimentos)}")

# Verificar se ambas as listas têm o mesmo comprimento
if len(categorias) != len(vencimentos):
    raise ValueError("As listas de categorias e vencimentos devem ter o mesmo tamanho!")

# Inicializando o DataFrame com as categorias, vencimentos e status
df = pd.DataFrame({
    'Categoria': categorias,
    'Vencimento': vencimentos,
    'Status': ['Pendente'] * len(categorias)  # Status inicial de todas as despesas
})

# Caminho do arquivo Excel onde a planilha será salva
file_path = "despesas.xlsx"

# Salvar o DataFrame no arquivo Excel
df.to_excel(file_path, index=False)

print(f"Planilha de despesas criada com sucesso: {file_path}")

df.loc[df['Categoria'] == 'Aluguel Casa', 'Status'] = 'Pago'
df.to_excel(file_path, index=False)  # Salvar novamente após alterações
