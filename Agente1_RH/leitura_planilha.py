from openpyxl import load_workbook

# Nome da planilha dentro da pasta do projeto.
nome_arquivo = "teste_coluna_ausente.xlsx"

# Lista inicial de colunas que queremos verificar.
colunas_obrigatorias = [
    "Número",
    "Status",
    "ANS Expirado",
    "Categoria",
    "Item",
    "Ação",
    "Requisitado para",
    "Grupo Designado",
    "Atribuído a",
    "Descrição",
    "Descrição Resumida",
    "Email do Solicitante",
    "Criação em",
    "Email From",
    "Integration ID",
    "Aberto por",
    "Duration Function Task",
    "Encerrado",
    "Aberto",
    "Anexo",
    "Anotações de Fechamento",
    "Anotações de Trabalho",
    "Aprovação",
    "Ativo(a)",
    "Atualização de",
    "Atualização em",
    "Atualizações",
    "Base de Conhecimento",
    "CC (cópia)",
    "Canal de Contato",
    "Comentários Adicionais",
    "Contagem de Reatribuições",
    "Criação de",
    "Encerrado por"
]

# Abre a planilha somente para leitura.
planilha = load_workbook(nome_arquivo, read_only=True)
aba = planilha.worksheets[0]

# Guarda os valores da primeira linha em uma lista.
colunas_encontradas = []

for celula in aba[1]:
    colunas_encontradas.append(celula.value)

print("Planilha:", nome_arquivo)
print("Aba:", aba.title)
print()

# Verifica cada coluna da lista.
print("RESULTADO DA VALIDAÇÃO")
print("----------------------")

colunas_ausentes = []

for coluna in colunas_obrigatorias:
    if coluna in colunas_encontradas:
        print(f"OK: coluna '{coluna}' encontrada.")
    else:
        print(f"ATENÇÃO: coluna '{coluna}' não encontrada.")
        colunas_ausentes.append(coluna)

print()

# Mostra o resultado geral.
if len(colunas_ausentes) == 0:
    print("Todas as colunas da lista foram encontradas.")
else:
    print("Colunas que precisam ser verificadas:")
    for coluna in colunas_ausentes:
        print("-", coluna)

planilha.close()