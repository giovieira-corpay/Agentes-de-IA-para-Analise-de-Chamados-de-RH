import spacy

# Carrega o modelo em português (versão grande, instalada manualmente).
modelo = spacy.load("pt_core_news_lg")

# Exemplo fictício de um chamado.
texto = (
    "Maria Souza entrou em contato com Carlos Oliveira "
    "para conversar sobre uma solicitação."
)

# Pede ao modelo para analisar o texto.
resultado = modelo(texto)

print("Nomes que o modelo identificou:")
print()

# Percorre as entidades identificadas no texto.
for entidade in resultado.ents:
    if entidade.label_ == "PER":
        print("-", entidade.text)