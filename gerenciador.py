#!/usr/bin/env python3
"""Tela local para filtrar, somar e editar os JSONs desta pasta.

Uso:
    python3 gerenciador.py
    python3 gerenciador.py --porta 9000 --sem-navegador
    python3 gerenciador.py --pasta /outra/pasta

Usa só a biblioteca padrão e só aceita conexões deste computador. Ao salvar,
a versão anterior vai para backups/<nome>/<data e hora>.json e pode ser
restaurada pela tela ("Histórico").
"""

import argparse
import datetime
import io
import json
import os
import re
import shutil
import sys
import tempfile
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit
from xml.sax.saxutils import escape

from ods_para_json import converter

TELA = Path(__file__).with_name("gerenciador.html")
LIMITE_UPLOAD = 50 * 1024 * 1024


class ErroDaApi(Exception):
    def __init__(self, status, mensagem):
        super().__init__(mensagem)
        self.status = status
        self.mensagem = mensagem


AREAS = ("bases_por_dia", "resultados_por_dia")
PASTA_BACKUPS = "backups"
NOME_BACKUP = re.compile(r"^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}-\d{6}\.json$")


def caminho_do_json(pasta, nome):
    """Aceita "arquivo.json" na pasta ou "<área>/arquivo.json" numa das AREAS, nada além disso."""
    partes = nome.split("/") if nome else []
    if len(partes) == 2 and partes[0] in AREAS:
        area, arquivo = partes
    elif len(partes) == 1:
        area, arquivo = "", partes[0]
    else:
        raise ErroDaApi(400, "Nome de arquivo inválido.")
    if not arquivo or arquivo.startswith(".") or Path(arquivo).name != arquivo or Path(arquivo).suffix.lower() != ".json":
        raise ErroDaApi(400, "Nome de arquivo inválido: use apenas um nome terminado em .json.")
    return pasta / area / arquivo


def pasta_de_backup(pasta, nome):
    # "bases_por_dia/2026-11-01.json" -> backups/bases_por_dia/2026-11-01/
    return pasta / PASTA_BACKUPS / Path(nome).with_suffix("")


def resumo_dos_dados(dados):
    """Total e quantidade de itens, para reconhecer uma versão no histórico."""
    if isinstance(dados, list):
        valores = [d.get("value") for d in dados if isinstance(d, dict)]
        numeros = [v for v in valores if isinstance(v, (int, float)) and not isinstance(v, bool)]
        total = sum(round(v * 100) for v in numeros) / 100 if numeros else None
        return {"itens": len(dados), "total": total}
    if isinstance(dados, dict):
        return {"itens": sum(len(v) for v in dados.values() if isinstance(v, list)), "total": None}
    return {"itens": None, "total": None}


def versao(caminho):
    # Vai como texto: um inteiro em nanossegundos perderia precisão no JavaScript.
    return str(caminho.stat().st_mtime_ns)


def gravar_json(caminho, dados, destino_backup):
    """Guarda a versão atual no histórico e troca o arquivo de uma vez só."""
    caminho.parent.mkdir(parents=True, exist_ok=True)
    if caminho.exists():
        destino_backup.mkdir(parents=True, exist_ok=True)
        carimbo = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S-%f")
        shutil.copy2(caminho, destino_backup / f"{carimbo}.json")
    conteudo = json.dumps(dados, ensure_ascii=False, indent=2) + "\n"
    descritor, temporario = tempfile.mkstemp(dir=caminho.parent, suffix=".tmp")
    try:
        with os.fdopen(descritor, "w", encoding="utf-8") as arquivo:
            arquivo.write(conteudo)
        if caminho.exists():
            shutil.copymode(caminho, temporario)
        os.replace(temporario, caminho)
    except BaseException:
        Path(temporario).unlink(missing_ok=True)
        raise


ESTILOS_XLSX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;R$ &quot;#,##0.00"/></numFmts>
<fonts count="3">
<font><sz val="11"/><name val="Arial"/></font>
<font><b/><sz val="11"/><name val="Arial"/></font>
<font><i/><sz val="9"/><color rgb="FF646B61"/><name val="Arial"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFECEEE9"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFE2F1EA"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left/><right/><top style="thin"><color auto="1"/></top><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="10">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="1" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="1" fontId="1" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="right"/></xf>
</cellXfs>
</styleSheet>
"""
TEXTO, CABECALHO, MOEDA, INTEIRO, TEXTO_SOMADO, MOEDA_SOMADO, INTEIRO_SOMADO, MOEDA_TOTAL, NOTA, CABECALHO_NUM = range(10)


def celula_texto(ref, texto, estilo):
    # Texto vai sempre como texto: um nome começando com "=" não vira fórmula no Excel.
    limpo = "".join(c for c in str(texto) if c >= " " or c in "\t\n")
    return f'<c r="{ref}" s="{estilo}" t="inlineStr"><is><t xml:space="preserve">{escape(limpo)}</t></is></c>'


def celula_numero(ref, valor, estilo):
    if valor is None:
        return f'<c r="{ref}" s="{estilo}"/>'
    return f'<c r="{ref}" s="{estilo}"><v>{valor!r}</v></c>'


def celula_formula(ref, formula, valor, estilo):
    return f'<c r="{ref}" s="{estilo}"><f>{formula}</f><v>{valor!r}</v></c>'


def numero_ou_vazio(valor, permite_vazio=True):
    if valor is None and permite_vazio:
        return None
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        raise ErroDaApi(400, "Os valores da exportação precisam ser números.")
    return valor


def montar_xlsx(linhas, nota):
    """Monta o .xlsx da prévia. "Novo valor" e os totais são fórmulas do Excel."""
    titulos = ["Empresa", "Atual", "Somado", "Novo valor", "Occupied"]
    xml = ['<row r="1">' + "".join(
        celula_texto(f"{c}1", t, CABECALHO if c == "A" else CABECALHO_NUM) for c, t in zip("ABCDE", titulos)
    ) + "</row>"]
    total_atual = total_somado = 0
    for numero, linha in enumerate(linhas, start=2):
        atual = numero_ou_vazio(linha.get("atual"), permite_vazio=False)
        somado = numero_ou_vazio(linha.get("somado"))
        occupied = numero_ou_vazio(linha.get("occupied"))
        total_atual += round(atual * 100)
        total_somado += round((somado or 0) * 100)
        novo = (round(atual * 100) + round((somado or 0) * 100)) / 100
        texto, moeda, inteiro = (TEXTO, MOEDA, INTEIRO) if somado is None else (TEXTO_SOMADO, MOEDA_SOMADO, INTEIRO_SOMADO)
        xml.append(
            f'<row r="{numero}">'
            + celula_texto(f"A{numero}", linha.get("empresa", ""), texto)
            + celula_numero(f"B{numero}", atual, moeda)
            + celula_numero(f"C{numero}", somado, moeda)
            + celula_formula(f"D{numero}", f"B{numero}+C{numero}", novo, moeda)
            + celula_numero(f"E{numero}", occupied, inteiro)
            + "</row>"
        )
    ultima = len(linhas) + 1
    t = ultima + 1
    xml.append(
        f'<row r="{t}">'
        + celula_texto(f"A{t}", "Total", MOEDA_TOTAL)
        + celula_formula(f"B{t}", f"SUM(B2:B{ultima})", total_atual / 100, MOEDA_TOTAL)
        + celula_formula(f"C{t}", f"SUM(C2:C{ultima})", total_somado / 100, MOEDA_TOTAL)
        + celula_formula(f"D{t}", f"SUM(D2:D{ultima})", (total_atual + total_somado) / 100, MOEDA_TOTAL)
        + f'<c r="E{t}" s="{MOEDA_TOTAL}"/></row>'
    )
    if nota:
        xml.append(f'<row r="{t + 2}">' + celula_texto(f"A{t + 2}", nota, NOTA) + "</row>")

    principal = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    relacoes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    cabecalho_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    planilha = (
        f'{cabecalho_xml}<worksheet xmlns="{principal}">'
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        '<cols><col min="1" max="1" width="48" customWidth="1"/><col min="2" max="4" width="17" customWidth="1"/><col min="5" max="5" width="11" customWidth="1"/></cols>'
        f'<sheetData>{"".join(xml)}</sheetData><autoFilter ref="A1:E{ultima}"/><pageSetup fitToWidth="1" fitToHeight="0"/></worksheet>'
    )
    arquivos = {
        "[Content_Types].xml": f'{cabecalho_xml}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        "_rels/.rels": f'{cabecalho_xml}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{relacoes}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'{cabecalho_xml}<workbook xmlns="{principal}" xmlns:r="{relacoes}">'
        '<sheets><sheet name="Base somada" sheetId="1" r:id="rId1"/></sheets>'
        f"<definedNames><definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"0\" hidden=\"1\">'Base somada'!$A$1:$E${ultima}</definedName></definedNames>"
        '<calcPr fullCalcOnLoad="1"/></workbook>',
        "xl/_rels/workbook.xml.rels": f'{cabecalho_xml}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{relacoes}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{relacoes}/styles" Target="styles.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": planilha,
        "xl/styles.xml": ESTILOS_XLSX,
    }
    saida = io.BytesIO()
    with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as pacote:
        for nome, conteudo in arquivos.items():
            pacote.writestr(nome, conteudo)
    return saida.getvalue()


class Handler(BaseHTTPRequestHandler):
    pasta: Path

    def log_message(self, *args):
        pass

    def responder(self, status, corpo, tipo="application/json; charset=utf-8", cabecalhos=None):
        self.send_response(status)
        self.send_header("Content-Type", tipo)
        for nome, valor in (cabecalhos or {}).items():
            self.send_header(nome, valor)
        self.send_header("Content-Length", str(len(corpo)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(corpo)

    def responder_json(self, status, dados):
        self.responder(status, json.dumps(dados, ensure_ascii=False).encode("utf-8"))

    def origem_confiavel(self):
        """Bloqueia outros sites abertos no navegador de chegarem nesta tela."""
        porta = self.server.server_address[1]
        permitidos = {f"127.0.0.1:{porta}", f"localhost:{porta}"}
        if self.headers.get("Host") not in permitidos:
            return False
        origem = self.headers.get("Origin")
        return origem is None or urlsplit(origem).netloc in permitidos

    def ler_corpo(self):
        try:
            tamanho = int(self.headers.get("Content-Length", ""))
        except ValueError:
            raise ErroDaApi(411, "Tamanho do envio não informado.")
        if tamanho > LIMITE_UPLOAD:
            raise ErroDaApi(413, "Arquivo grande demais (limite de 50 MB).")
        return self.rfile.read(tamanho)

    def tratar(self, acao):
        if not self.origem_confiavel():
            self.responder_json(403, {"erro": "Origem não permitida."})
            return
        try:
            acao(urlsplit(self.path))
        except ErroDaApi as erro:
            self.responder_json(erro.status, {"erro": erro.mensagem})
        except OSError as erro:
            self.responder_json(500, {"erro": f"Erro ao acessar o arquivo: {erro}"})

    def do_GET(self):
        self.tratar(self.rota_get)

    def do_POST(self):
        self.tratar(self.rota_post)

    def rota_get(self, url):
        if url.path == "/":
            self.responder(200, TELA.read_bytes(), "text/html; charset=utf-8")
        elif url.path == "/api/arquivos":
            area = parse_qs(url.query).get("area", [""])[0]
            if area and area not in AREAS:
                raise ErroDaApi(400, "Área inválida.")
            prefixo = f"{area}/" if area else ""
            nomes = sorted(prefixo + c.name for c in (self.pasta / area).glob("*.json") if c.is_file())
            self.responder_json(200, {"arquivos": nomes})
        elif url.path == "/api/historico":
            self.historico(parse_qs(url.query).get("nome", [""])[0])
        elif url.path == "/api/backup.zip":
            self.baixar_backup()
        elif url.path == "/api/arquivo":
            caminho = caminho_do_json(self.pasta, parse_qs(url.query).get("nome", [""])[0])
            if not caminho.is_file():
                raise ErroDaApi(404, "Arquivo não encontrado.")
            try:
                dados = json.loads(caminho.read_text(encoding="utf-8"))
            except ValueError as erro:
                raise ErroDaApi(422, f"O arquivo não é um JSON válido: {erro}")
            self.responder_json(200, {"dados": dados, "versao": versao(caminho)})
        else:
            raise ErroDaApi(404, "Rota não encontrada.")

    def rota_post(self, url):
        if url.path == "/api/salvar":
            self.salvar()
        elif url.path == "/api/restaurar":
            self.restaurar()
        elif url.path == "/api/exportar-xlsx":
            self.exportar_xlsx()
        elif url.path == "/api/importar":
            self.importar(parse_qs(url.query).get("nome", ["planilha.ods"])[0])
        else:
            raise ErroDaApi(404, "Rota não encontrada.")

    def salvar(self):
        try:
            pedido = json.loads(self.ler_corpo())
            nome, dados = pedido["nome"], pedido["dados"]
        except (ValueError, KeyError, TypeError):
            raise ErroDaApi(400, "Pedido de salvamento inválido.")
        caminho = caminho_do_json(self.pasta, nome)
        if not pedido.get("forcar"):
            atual = versao(caminho) if caminho.exists() else None
            if atual != pedido.get("versao"):
                raise ErroDaApi(409, "O arquivo no disco é diferente do que está aberto aqui.")
        gravar_json(caminho, dados, pasta_de_backup(self.pasta, nome))
        self.responder_json(200, {"versao": versao(caminho)})

    def historico(self, nome):
        caminho_do_json(self.pasta, nome)
        pasta = pasta_de_backup(self.pasta, nome)
        versoes = []
        if pasta.is_dir():
            for arquivo in sorted(pasta.iterdir(), reverse=True):
                if not NOME_BACKUP.match(arquivo.name):
                    continue
                try:
                    dados = json.loads(arquivo.read_text(encoding="utf-8"))
                except ValueError:
                    dados = None
                versoes.append({"id": arquivo.name, **resumo_dos_dados(dados)})
        self.responder_json(200, {"versoes": versoes})

    def restaurar(self):
        try:
            pedido = json.loads(self.ler_corpo())
            nome, versao_id = pedido["nome"], pedido["id"]
        except (ValueError, KeyError, TypeError):
            raise ErroDaApi(400, "Pedido de restauração inválido.")
        caminho = caminho_do_json(self.pasta, nome)
        if not isinstance(versao_id, str) or not NOME_BACKUP.match(versao_id):
            raise ErroDaApi(400, "Versão inválida.")
        origem = pasta_de_backup(self.pasta, nome) / versao_id
        if not origem.is_file():
            raise ErroDaApi(404, "Essa versão não existe mais.")
        try:
            dados = json.loads(origem.read_text(encoding="utf-8"))
        except ValueError:
            raise ErroDaApi(422, "Essa versão do histórico está corrompida.")
        # A versão atual também vai para o histórico, então restaurar pode ser desfeito.
        gravar_json(caminho, dados, pasta_de_backup(self.pasta, nome))
        self.responder_json(200, {"versao": versao(caminho), "dados": dados})

    def baixar_backup(self):
        """Um .zip com os JSONs, as bases e resultados por dia e todo o histórico."""
        saida = io.BytesIO()
        with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as pacote:
            arquivos = [*self.pasta.glob("*.json"), *(a for area in AREAS for a in (self.pasta / area).glob("*.json")),
                        *(self.pasta / PASTA_BACKUPS).rglob("*.json")]
            for arquivo in arquivos:
                if arquivo.is_file():
                    pacote.write(arquivo, arquivo.relative_to(self.pasta).as_posix())
        nome = f"backup_planilhas_{datetime.datetime.now():%Y-%m-%d_%H-%M}.zip"
        self.responder(200, saida.getvalue(), "application/zip", {"Content-Disposition": f'attachment; filename="{nome}"'})

    def exportar_xlsx(self):
        try:
            pedido = json.loads(self.ler_corpo())
            linhas = pedido["linhas"]
            if not isinstance(linhas, list) or not all(isinstance(l, dict) for l in linhas):
                raise ValueError
        except (ValueError, KeyError, TypeError):
            raise ErroDaApi(400, "Pedido de exportação inválido.")
        tipo = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        self.responder(200, montar_xlsx(linhas, pedido.get("nota")), tipo)

    def importar(self, nome):
        try:
            dados = converter(io.BytesIO(self.ler_corpo()))
        except (KeyError, ValueError, zipfile.BadZipFile, ET.ParseError) as erro:
            raise ErroDaApi(400, f"Não foi possível ler a planilha: {erro}")
        self.responder_json(200, {"dados": dados, "nome": f"{Path(Path(nome).name).stem}.json"})


def main():
    parser = argparse.ArgumentParser(description="Tela para filtrar, somar e editar os JSONs.")
    parser.add_argument("--porta", type=int, default=8765)
    parser.add_argument("--pasta", type=Path, default=Path(__file__).parent, help="pasta dos arquivos JSON")
    parser.add_argument("--sem-navegador", action="store_true", help="não abre o navegador")
    args = parser.parse_args()

    if not args.pasta.is_dir():
        sys.exit(f"Pasta não encontrada: {args.pasta}")
    Handler.pasta = args.pasta.resolve()

    try:
        servidor = ThreadingHTTPServer(("127.0.0.1", args.porta), Handler)
    except OSError as erro:
        sys.exit(f"Não foi possível usar a porta {args.porta}: {erro}. Tente --porta 9000.")

    endereco = f"http://127.0.0.1:{args.porta}/"
    print(f"Tela aberta em {endereco}  (Ctrl+C para encerrar)")
    if not args.sem_navegador:
        webbrowser.open(endereco)
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print()
    finally:
        servidor.server_close()


if __name__ == "__main__":
    main()
