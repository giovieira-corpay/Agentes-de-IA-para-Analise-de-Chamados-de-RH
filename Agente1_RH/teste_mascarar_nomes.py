import spacy

# Carrega o modelo de português instalado.
modelo = spacy.load("pt_core_news_lg")


# Lista inicial para testar a identificação de conteúdo sensível.
palavras_sensiveis = [
    "assédio",
    "violência",
    "discriminação",
    "doença",
    "tratamento médico",
    "deficiência",
]


def contem_conteudo_sensivel(texto):
    # Converte o texto para letras minúsculas na comparação.
    texto_minusculo = texto.lower()

    for palavra in palavras_sensiveis:
        if palavra in texto_minusculo:
            return True

    return False


def mascarar_nomes(texto):
    resultado = modelo(texto)
    texto_mascarado = texto

    numero_pessoa = 1

    # Substituímos de trás para frente para preservar as posições.
    for entidade in reversed(resultado.ents):
        if entidade.label_ == "PER":
            codigo = f"PESSOA_{numero_pessoa:04d}"

            texto_mascarado = (
                texto_mascarado[:entidade.start_char]
                + codigo
                + texto_mascarado[entidade.end_char:]
            )

            numero_pessoa += 1

    return texto_mascarado


def processar_texto(texto):
    if contem_conteudo_sensivel(texto):
        print("Sensível: Sim")
        return mascarar_nomes(texto)

    print("Sensível: Não identificado no teste")
    return texto


# --------------------------------------------------
# TESTES COM DADOS FICTÍCIOS
# --------------------------------------------------

exemplos = [
    "Maria Souza entrou em contato para atualizar seu cadastro.",
    "Maria Souza entrou em contato para relatar assédio.",
]

for texto in exemplos:
    print()
    print("TEXTO ORIGINAL:")
    print(texto)

    print("RESULTADO:")
    print(processar_texto(texto))