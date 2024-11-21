import pandas as pd

dados = {
    'Nome': ['Ana', 'Carlos', 'Bianca', 'João'],
    'Idade': [28, 34, 22, 41],
    'Cidade': ['São Paulo', 'Rio de Janeiro', 'Belo Horizonte', 'Curitiba'],
    'Remuneração': [2700, 3500, 2700, 4200]
}


df = pd.DataFrame.from_dict(dados)
df.to_excel("dados.xlsx", index=False)




