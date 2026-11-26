import ast
import re
import unicodedata

from pathlib import Path
from uuid import uuid4

import spacy
from openpyxl import load_workbook


# ==========================================================
# 1. ARQUIVOS E COLUNAS
# ==========================================================

# O programa procura os arquivos na mesma pasta deste código.
PASTA_PROJETO = Path(__file__).resolve().parent

ARQUIVO_CHAMADOS = (
    PASTA_PROJETO
    / "relatorio_semanal_teste4.xlsx"
)

ARQUIVO_BOT = (
    PASTA_PROJETO
    / "MeuRH_Interacoes_Ficticias_Realistas_Sensiveis.xlsx"
)

ARQUIVO_PALAVRAS = (
    PASTA_PROJETO
    / "palavras-chave-mascaramento.txt"
)

SAIDA_CHAMADOS = (
    PASTA_PROJETO
    / "chamados_mascarados_REVISAO.xlsx"
)

SAIDA_BOT = (
    PASTA_PROJETO
    / "bot_mascarado_REVISAO.xlsx"
)

COLUNAS_CHAMADOS = [
    "Descrição",
    "Anotações de Trabalho",
    "Anotações de Fechamento",
    "Comentários Adicionais",
]

COLUNAS_BOT = [
    "short_description",
    "generated_response",
    "user_input",
]


# Guarda os códigos de e-mail durante a execução atual.
# A mesma lista será usada nas duas planilhas.
emails_encontrados = {}


# ==========================================================
# 2. LEITURA DAS PALAVRAS SENSÍVEIS
# ==========================================================

def carregar_palavras_sensiveis(caminho):
    """
    Lê a lista chamada palavras_sensiveis do arquivo TXT.
    """

    conteudo = caminho.read_text(encoding="utf-8-sig")

    try:
        estrutura = ast.parse(conteudo)
    except SyntaxError as erro:
        raise ValueError(
            "Não foi possível ler o arquivo de palavras-chave. "
            "Confira se a lista está completa."
        ) from erro

    for item in estrutura.body:
        if isinstance(item, ast.Assign):
            for destino in item.targets:
                if (
                    isinstance(destino, ast.Name)
                    and destino.id == "palavras_sensiveis"
                ):
                    lista = ast.literal_eval(item.value)

                    if not isinstance(lista, list):
                        raise ValueError(
                            "'palavras_sensiveis' precisa ser uma lista."
                        )

                    if not all(
                        isinstance(palavra, str) and palavra.strip()
                        for palavra in lista
                    ):
                        raise ValueError(
                            "Cada palavra da lista precisa ser um texto "
                            "não vazio, escrito entre aspas."
                        )

                    return lista

    raise ValueError(
        "A lista 'palavras_sensiveis' não foi encontrada no TXT."
    )


def normalizar_para_comparacao(texto):
    """
    Remove diferenças de maiúsculas e acentos na comparação.
    """

    texto = texto.casefold()
    texto = unicodedata.normalize("NFKD", texto)

    return "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )


def preparar_padroes_sensiveis(palavras):
    """
    Prepara as palavras e frases para busca nos textos.
    """

    padroes = []

    for palavra in palavras:
        palavra_normalizada = normalizar_para_comparacao(
            palavra.strip()
        )

        padrao = re.compile(
            r"(?<!\w)"
            + re.escape(palavra_normalizada)
            + r"(?!\w)"
        )

        padroes.append(padrao)

    return padroes


def contem_palavra_sensivel(textos, padroes):
    """
    Procura palavras sensíveis no conjunto de textos da linha.
    """

    for texto in textos:
        if not isinstance(texto, str):
            continue

        texto_normalizado = normalizar_para_comparacao(texto)

        for padrao in padroes:
            if padrao.search(texto_normalizado):
                return True

    return False


# ==========================================================
# 3. MASCARAMENTO DOS DADOS PESSOAIS
# ==========================================================

def mascarar_emails(texto):
    """
    Exemplo:
    maria.silva@empresa.com.br -> USER_ID_0001
    """

    padrao = r"[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}"

    def substituir(resultado):
        email = resultado.group()
        email_comparacao = email.casefold()

        if email_comparacao not in emails_encontrados:
            numero = len(emails_encontrados) + 1

            emails_encontrados[email_comparacao] = (
                f"USER_ID_{numero:04d}"
            )

        return emails_encontrados[email_comparacao]

    return re.sub(padrao, substituir, texto)


def mascarar_cpf(texto):
    """
    Reconhece o CPF pontuado e o CPF sem pontuação
    quando ele está identificado pela palavra CPF.
    """

    padrao_formatado = (
        r"(?<!\d)(\d{3})\.(\d{3})\.(\d{3})-(\d{2})(?!\d)"
    )

    def substituir_formatado(resultado):
        inicio = resultado.group(1)
        final = resultado.group(4)

        return f"{inicio}.***.***-{final}"

    texto = re.sub(
        padrao_formatado,
        substituir_formatado,
        texto
    )

    padrao_identificado = (
        r"(?i)(\bCPF\s*[:\-]?\s*)(\d{3})(\d{6})(\d{2})(?!\d)"
    )

    def substituir_identificado(resultado):
        identificacao = resultado.group(1)
        inicio = resultado.group(2)
        final = resultado.group(4)

        return f"{identificacao}{inicio}.***.***-{final}"

    return re.sub(
        padrao_identificado,
        substituir_identificado,
        texto
    )


def mascarar_matriculas(texto):
    """
    Exemplo:
    matrícula: 12345678 -> matrícula: ******78
    """

    padrao = (
        r"(?i)(\bmatr[ií]cula\s*[:\-]?\s*)(\d{3,})(?!\d)"
    )

    def substituir(resultado):
        identificacao = resultado.group(1)
        numero = resultado.group(2)

        parte_oculta = "*" * (len(numero) - 2)
        final = numero[-2:]

        return identificacao + parte_oculta + final

    return re.sub(padrao, substituir, texto)


def mascarar_telefones(texto):
    """
    Exemplos:
    (11) 99888-7777 -> (11) *****-7777
    (11) 3333-4444  -> (11) ****-4444
    """

    padrao = (
        r"(?<!\d)(\(\d{2}\)\s*)(\d{4,5})(-)(\d{4})(?!\d)"
    )

    def substituir(resultado):
        ddd = resultado.group(1)
        parte_central = resultado.group(2)
        separador = resultado.group(3)
        numeros_finais = resultado.group(4)

        return (
            ddd
            + "*" * len(parte_central)
            + separador
            + numeros_finais
        )

    return re.sub(padrao, substituir, texto)


def mascarar_dados_pessoais(texto):
    """
    Aplica as quatro regras de mascaramento ao texto.
    """

    if not isinstance(texto, str):
        return texto

    texto = mascarar_emails(texto)
    texto = mascarar_cpf(texto)
    texto = mascarar_matriculas(texto)
    texto = mascarar_telefones(texto)

    return texto


# ==========================================================
# 4. IDENTIFICAÇÃO E MASCARAMENTO DE NOMES
# ==========================================================

def identificar_nomes(textos, modelo):
    """
    Tenta identificar nomes nas colunas analisadas da linha.
    """

    nomes = {}

    for texto in textos:
        if not isinstance(texto, str) or not texto.strip():
            continue

        resultado = modelo(texto)

        for entidade in resultado.ents:
            if entidade.label_ == "PER":
                nome = entidade.text.strip()
                chave = nome.casefold()

                if nome and chave not in nomes:
                    nomes[chave] = nome

    return list(nomes.values())


def criar_codigos_nomes(nomes):
    """
    Cria um código aleatório para cada nome identificado.

    Os códigos são consistentes entre as colunas da mesma
    linha quando o nome aparece com a mesma escrita.
    """

    codigos = {}

    for nome in nomes:
        codigo = "PESSOA_" + uuid4().hex[:12].upper()
        codigos[nome] = codigo

    return codigos


def mascarar_nomes_no_texto(texto, codigos):
    """
    Substitui no texto os nomes identificados pelo spaCy.
    """

    if not isinstance(texto, str):
        return texto

    # Trata os nomes maiores primeiro.
    nomes_ordenados = sorted(
        codigos.keys(),
        key=len,
        reverse=True
    )

    for nome in nomes_ordenados:
        padrao = (
            r"(?<!\w)"
            + re.escape(nome)
            + r"(?!\w)"
        )

        texto = re.sub(
            padrao,
            lambda resultado: codigos[nome],
            texto,
            flags=re.IGNORECASE
        )

    return texto


# ==========================================================
# 5. PROCESSAMENTO DE UMA PLANILHA
# ==========================================================

def processar_planilha(
    arquivo_entrada,
    arquivo_saida,
    colunas_analisadas,
    nome_aba,
    modelo,
    padroes_sensiveis
):
    """
    Trata as colunas informadas e salva uma cópia.

    As outras colunas e abas são mantidas sem mascaramento.
    """

    print()
    print(f"Processando: {arquivo_entrada.name}")

    planilha = load_workbook(arquivo_entrada)

    try:
        # Nos chamados, usamos a primeira aba.
        # No bot, usamos a aba chamada Interações.
        if nome_aba is None:
            aba = planilha.worksheets[0]
        else:
            if nome_aba not in planilha.sheetnames:
                raise ValueError(
                    f"A aba '{nome_aba}' não foi encontrada "
                    f"em {arquivo_entrada.name}."
                )

            aba = planilha[nome_aba]

        # Os nomes das colunas devem estar na primeira linha.
        cabecalhos = [
            celula.value
            for celula in aba[1]
        ]

        nomes_presentes = [
            nome
            for nome in cabecalhos
            if nome is not None
        ]

        repetidos = sorted({
            nome
            for nome in nomes_presentes
            if nomes_presentes.count(nome) > 1
        })

        if repetidos:
            raise ValueError(
                f"No arquivo {arquivo_entrada.name}, "
                "há colunas com nomes repetidos: "
                + ", ".join(str(nome) for nome in repetidos)
            )

        ausentes = [
            coluna
            for coluna in colunas_analisadas
            if coluna not in cabecalhos
        ]

        if ausentes:
            raise ValueError(
                f"No arquivo {arquivo_entrada.name}, "
                "as seguintes colunas não foram encontradas: "
                + ", ".join(ausentes)
            )

        posicoes = {
            nome: cabecalhos.index(nome) + 1
            for nome in colunas_analisadas
        }

        total_linhas = 0
        total_sinalizadas = 0

        # A primeira linha tem os títulos das colunas.
        for numero_linha in range(2, aba.max_row + 1):
            valores = {
                coluna: aba.cell(
                    row=numero_linha,
                    column=posicoes[coluna]
                ).value
                for coluna in colunas_analisadas
            }

            textos_originais = list(valores.values())

            # Verifica as palavras sensíveis no texto original.
            sensivel = contem_palavra_sensivel(
                textos_originais,
                padroes_sensiveis
            )

            # Dados pessoais são tratados em todas as linhas.
            textos_tratados = {
                coluna: mascarar_dados_pessoais(valor)
                for coluna, valor in valores.items()
            }

            codigos_nomes = {}

            if sensivel:
                total_sinalizadas += 1

                # Procura nomes depois de mascarar os dados pessoais,
                # para não confundir parte de um e-mail com um nome.
                nomes = identificar_nomes(
                    list(textos_tratados.values()),
                    modelo
                )

                codigos_nomes = criar_codigos_nomes(nomes)

            for coluna in colunas_analisadas:
                celula = aba.cell(
                    row=numero_linha,
                    column=posicoes[coluna]
                )

                valor_original = valores[coluna]

                # Preserva células vazias e valores não textuais.
                if not isinstance(valor_original, str):
                    continue

                # Interrompe caso encontre uma fórmula nas colunas
                # que deveriam conter texto livre.
                if celula.data_type == "f":
                    raise ValueError(
                        f"Foi encontrada uma fórmula no arquivo "
                        f"{arquivo_entrada.name}, linha {numero_linha}, "
                        f"coluna '{coluna}'. Revise antes de continuar."
                    )

                texto_tratado = textos_tratados[coluna]

                if sensivel:
                    texto_tratado = mascarar_nomes_no_texto(
                        texto_tratado,
                        codigos_nomes
                    )

                celula.value = texto_tratado

            total_linhas += 1

        # Cria a saída somente se ela ainda não existir.
        with open(arquivo_saida, "xb") as saida:
            planilha.save(saida)

        print(f"Aba processada: {aba.title}")
        print(f"Linhas processadas: {total_linhas}")
        print(f"Linhas com palavra sensível: {total_sinalizadas}")
        print(f"Arquivo criado: {arquivo_saida.name}")

    finally:
        planilha.close()


# ==========================================================
# 6. EXECUÇÃO DAS DUAS PLANILHAS
# ==========================================================

def main():
    print("Iniciando o processamento...")

    arquivos_necessarios = [
        ARQUIVO_CHAMADOS,
        ARQUIVO_BOT,
        ARQUIVO_PALAVRAS,
    ]

    for arquivo in arquivos_necessarios:
        if not arquivo.is_file():
            raise FileNotFoundError(
                f"Arquivo não encontrado: {arquivo.name}"
            )

    # Evita substituir resultados anteriores.
    for arquivo in [SAIDA_CHAMADOS, SAIDA_BOT]:
        if arquivo.exists():
            raise FileExistsError(
                f"Já existe um arquivo chamado {arquivo.name}. "
                "Renomeie o resultado anterior antes de continuar."
            )

    palavras = carregar_palavras_sensiveis(
        ARQUIVO_PALAVRAS
    )

    padroes_sensiveis = preparar_padroes_sensiveis(
        palavras
    )

    print(f"Palavras e frases carregadas: {len(palavras)}")
    print("Carregando o modelo do spaCy...")

    modelo = spacy.load("pt_core_news_lg")

    # Processa o relatório de chamados.
    processar_planilha(
        arquivo_entrada=ARQUIVO_CHAMADOS,
        arquivo_saida=SAIDA_CHAMADOS,
        colunas_analisadas=COLUNAS_CHAMADOS,
        nome_aba=None,
        modelo=modelo,
        padroes_sensiveis=padroes_sensiveis
    )

    # Processa a aba Interações do bot Meu RH.
    processar_planilha(
        arquivo_entrada=ARQUIVO_BOT,
        arquivo_saida=SAIDA_BOT,
        colunas_analisadas=COLUNAS_BOT,
        nome_aba="Interações",
        modelo=modelo,
        padroes_sensiveis=padroes_sensiveis
    )

    print()
    print("Processamento das duas planilhas concluído.")
    print("Revise as saídas antes de qualquer utilização.")


if __name__ == "__main__":
    main()