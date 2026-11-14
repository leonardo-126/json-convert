#!/usr/bin/env python3
"""Converte uma planilha .ods em JSON. Usa só a biblioteca padrão do Python.

Uso:
    python3 ods_para_json.py "planilha.ods"
    python3 ods_para_json.py "planilha.ods" -o saida.json
    python3 ods_para_json.py "planilha.ods" --aba "Planilha1"
    python3 ods_para_json.py "planilha.ods" --manter-cabecalhos

Regras:
- A primeira linha preenchida de cada aba é o cabeçalho.
- Cabeçalhos viram snake_case ("VALOR R$ " -> "valor_r"), a menos que se use
  --manter-cabecalhos.
- Números, moedas e porcentagens saem como número (2563.96); datas em ISO
  (2026-10-19); células vazias como null.
- Uma aba gera uma lista de objetos; várias abas geram {"nome da aba": [...]}.
"""

import argparse
import json
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
TEXT = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
CELLS = (TABLE + "table-cell", TABLE + "covered-table-cell")


def texto_do_no(no):
    """Extrai o texto de um nó, incluindo os espaços que o ODS guarda como tags."""
    partes = [no.text or ""]
    for filho in no:
        if filho.tag == TEXT + "s":
            partes.append(" " * int(filho.get(TEXT + "c", "1")))
        elif filho.tag == TEXT + "tab":
            partes.append("\t")
        elif filho.tag == TEXT + "line-break":
            partes.append("\n")
        else:
            partes.append(texto_do_no(filho))
        partes.append(filho.tail or "")
    return "".join(partes)


def valor_da_celula(celula):
    """Devolve o valor real da célula, não o texto formatado que aparece na tela."""
    tipo = celula.get(OFFICE + "value-type")

    if tipo in ("float", "currency", "percentage"):
        numero = float(celula.get(OFFICE + "value"))
        return int(numero) if numero.is_integer() else numero
    if tipo == "date":
        return celula.get(OFFICE + "date-value").removesuffix("T00:00:00")
    if tipo == "time":
        return celula.get(OFFICE + "time-value")
    if tipo == "boolean":
        return celula.get(OFFICE + "boolean-value") == "true"

    texto = "\n".join(texto_do_no(p) for p in celula.findall(TEXT + "p")).strip()
    return texto or None


def ler_linha(linha):
    """Lê uma linha sem as células vazias à direita.

    O LibreOffice marca as colunas sem uso com um "repeat" de milhares de células;
    por isso os vazios só são expandidos quando há um valor depois deles.
    """
    valores = []
    vazios_pendentes = 0
    for celula in linha:
        if celula.tag not in CELLS:
            continue
        repeticoes = int(celula.get(TABLE + "number-columns-repeated", "1"))
        valor = valor_da_celula(celula)
        if valor is None:
            vazios_pendentes += repeticoes
            continue
        valores.extend([None] * vazios_pendentes)
        vazios_pendentes = 0
        valores.extend([valor] * repeticoes)
    return valores


def ler_linhas(aba):
    """Devolve as linhas preenchidas da aba, ignorando as totalmente vazias."""
    linhas = []
    for linha in aba.iter(TABLE + "table-row"):
        valores = ler_linha(linha)
        if valores:
            repeticoes = int(linha.get(TABLE + "number-rows-repeated", "1"))
            linhas.extend([valores] * repeticoes)
    return linhas


def snake_case(texto):
    sem_acento = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", sem_acento.lower()).strip("_")


def montar_chaves(cabecalho, total_colunas, manter_cabecalhos):
    """Cria chaves únicas: coluna sem título vira coluna_N e repetida ganha _2, _3..."""
    chaves = []
    usadas = {}
    for indice in range(total_colunas):
        titulo = cabecalho[indice] if indice < len(cabecalho) else None
        titulo = str(titulo).strip() if titulo is not None else ""
        chave = titulo if manter_cabecalhos else snake_case(titulo)
        chave = chave or f"coluna_{indice + 1}"
        usadas[chave] = usadas.get(chave, 0) + 1
        chaves.append(chave if usadas[chave] == 1 else f"{chave}_{usadas[chave]}")
    return chaves


def aba_para_registros(aba, manter_cabecalhos):
    linhas = ler_linhas(aba)
    if not linhas:
        return []
    cabecalho, *dados = linhas
    total_colunas = max(len(linha) for linha in linhas)
    chaves = montar_chaves(cabecalho, total_colunas, manter_cabecalhos)
    return [
        {chave: (linha[i] if i < len(linha) else None) for i, chave in enumerate(chaves)}
        for linha in dados
    ]


def converter(caminho, nome_aba=None, manter_cabecalhos=False):
    with zipfile.ZipFile(caminho) as arquivo:
        raiz = ET.fromstring(arquivo.read("content.xml"))

    abas = {aba.get(TABLE + "name"): aba for aba in raiz.iter(TABLE + "table")}
    if nome_aba is not None:
        if nome_aba not in abas:
            disponiveis = ", ".join(abas)
            raise ValueError(f'Aba "{nome_aba}" não encontrada. Disponíveis: {disponiveis}')
        return aba_para_registros(abas[nome_aba], manter_cabecalhos)

    resultado = {nome: aba_para_registros(aba, manter_cabecalhos) for nome, aba in abas.items()}
    return next(iter(resultado.values())) if len(resultado) == 1 else resultado


def main():
    parser = argparse.ArgumentParser(description="Converte uma planilha .ods em JSON.")
    parser.add_argument("entrada", type=Path, help="caminho do arquivo .ods")
    parser.add_argument("-o", "--saida", type=Path, help="arquivo JSON (padrão: nome da planilha nesta pasta)")
    parser.add_argument("--aba", help="converte só esta aba")
    parser.add_argument("--manter-cabecalhos", action="store_true", help="não converte os cabeçalhos para snake_case")
    args = parser.parse_args()

    saida = args.saida or Path(__file__).parent / f"{args.entrada.stem}.json"

    try:
        dados = converter(args.entrada, args.aba, args.manter_cabecalhos)
    except (OSError, KeyError, ValueError, zipfile.BadZipFile, ET.ParseError) as erro:
        sys.exit(f"Erro ao ler {args.entrada}: {erro}")

    saida.write_text(json.dumps(dados, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    total = len(dados) if isinstance(dados, list) else sum(len(v) for v in dados.values())
    print(f"{total} registros gravados em {saida}")


if __name__ == "__main__":
    main()
