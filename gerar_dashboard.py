"""Gera dashboard_slim.json e injeta os dados no dashboard.html a partir da
planilha de controle (dados/*.xlsx).

Uso:
    python gerar_dashboard.py            # usa a planilha mais recente em dados/
    python gerar_dashboard.py --check    # só valida, sem gravar nada

O mapeamento de colunas foi construído lendo a estrutura da planilha-modelo
"CONTROLE - CLIENTES - BIOMASSA CAVACO - TERRA BIOMASSA" (abas EMBARQUES,
CONTRATOS, POSIÇÃO POR CONTRATO, TRANSPORTADORAS, MOTORISTAS, AUDITORIA,
CLIENTES E LOCAIS, PAINEL, CAPA) e validado célula a célula contra o
dashboard_slim.json já existente. Se a estrutura de colunas da planilha for
alterada no futuro, rode com --check para comparar o resultado contra o
dashboard_slim.json anterior antes de publicar.
"""

import argparse
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl

BASE_DIR = Path(__file__).resolve().parent
DADOS_DIR = BASE_DIR / "dados"
JSON_PATH = BASE_DIR / "dashboard_slim.json"
HTML_SOURCE = BASE_DIR / "dashboard.html"
HTML_PUBLISH = BASE_DIR / "index.html"

def achar_planilha() -> Path:
    arquivos = sorted(DADOS_DIR.glob("*.xlsx"))
    if not arquivos:
        raise SystemExit(f"Nenhuma planilha .xlsx encontrada em {DADOS_DIR}")
    return arquivos[-1]


def num(v, casas=2):
    if v is None or isinstance(v, str):
        return v
    r = round(float(v), casas)
    return int(r) if r == int(r) else r


def txt(v):
    return v.strip() if isinstance(v, str) else v


def data_iso(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return v


def limpa_header(texto):
    return re.sub(r"\s+", " ", str(texto)).strip()


def linhas(ws, header_row, key_col=1):
    """Uma linha por embarque/contrato/etc — pula linhas com a coluna-chave vazia,
    mas continua até o fim da aba (usado nas abas de tabela única)."""
    for r in range(header_row + 1, ws.max_row + 1):
        if ws.cell(row=r, column=key_col).value is not None:
            yield r


def linhas_ate_branco(ws, header_row, ultima_col):
    """Usado nas abas com duas tabelas (TRANSPORTADORAS, MOTORISTAS): para na
    primeira linha totalmente vazia, para não invadir a tabela seguinte."""
    for r in range(header_row + 1, ws.max_row + 1):
        valores = [ws.cell(row=r, column=c).value for c in range(1, ultima_col + 1)]
        if all(x is None for x in valores):
            return
        yield r


def gerar(xlsx_path: Path) -> dict:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # ---------------- META ----------------
    emissao_raw = wb["CAPA"]["C24"].value
    if isinstance(emissao_raw, str):
        emissao = datetime.strptime(emissao_raw.strip(), "%d/%m/%Y").strftime("%Y-%m-%d")
    else:
        emissao = data_iso(emissao_raw)

    painel = wb["PAINEL"]
    meta = {
        "empresa": "TERRA BIOMASSA INDUSTRIA E COMERCIO LTDA",
        "produto": "BIOMASSA - CAVACO DE MADEIRA",
        "emissao": emissao,
        "periodoPainelDe": data_iso(painel["B7"].value),
        "periodoPainelAte": data_iso(painel["B8"].value),
    }

    # ---------------- EMBARQUES ----------------
    ws = wb["EMBARQUES"]
    embarques = []
    for r in linhas(ws, header_row=6):
        c = lambda col: txt(ws.cell(row=r, column=col).value)
        pendencias = num(c(89), 0)
        embarques.append({
            "num": num(c(1), 0),
            "contrato": c(2),
            "cliente": c(3),
            "local": c(4),
            "transportadora": c(9),
            "motorista": c(12),
            "dataCarreg": data_iso(c(14)),
            "dataDescarga": data_iso(c(42)),
            "volProg": num(c(8)),
            "volEmb": num(c(16)),
            "volNf": num(c(24)),
            "volDescarregado": num(c(44), 3),
            "volLiquido": num(c(47)),
            "difCubM3": num(c(48)),
            "difCubPct": num(c(49), 4),
            "volComplemento": num(c(55)),
            "valorNf": num(c(26)),
            "valorComplemento": num(c(57)),
            "receitaTotal": num(c(63)),
            "valorAFaturar": num(c(62)),
            "statusConciliacao": c(64),
            "statusDescarga": c(51),
            "statusCiclo": c(88),
            "freteApurado": num(c(69)),
            "valorCte": num(c(37)),
            "aPagarTransp": num(c(70)),
            "valorPagoTransp": num(c(72)),
            "statusPgtoTransp": c(73),
            "creditoMotorista": num(c(75)),
            "saldoMotorista": num(c(84)),
            "statusPgtoMotorista": c(87),
            "pendencias": pendencias,
            "ocorrencias": pendencias,
            "precoContratado": num(c(6)),
        })

    # ---------------- CONTRATOS ----------------
    ws = wb["CONTRATOS"]
    contratos = []
    for r in linhas(ws, header_row=6):
        c = lambda col: txt(ws.cell(row=r, column=col).value)
        contratos.append({
            "contrato": c(1),
            "cliente": c(2),
            "produto": c(4),
            "volumeContratado": num(c(5)),
            "preco": num(c(6)),
            "inicioVigencia": data_iso(c(7)),
            "fimVigencia": data_iso(c(8)),
            "toleranciaPct": num(c(10)),
            "status": c(13),
            "condicaoPagamento": c(11),
        })

    # ---------------- POSIÇÃO POR CONTRATO ----------------
    ws = wb["POSIÇÃO POR CONTRATO"]
    posicao_contrato = []
    for r in linhas(ws, header_row=6):
        c = lambda col: txt(ws.cell(row=r, column=col).value)
        posicao_contrato.append({
            "contrato": c(1),
            "cliente": c(2),
            "volumeContratado": num(c(4)),
            "volumeLiquido": num(c(10)),
            "saldoEntregar": num(c(11)),
            "pctExecutado": num(c(12), 4),
            "receitaFaturada": num(c(13)),
            "valorAFaturar": num(c(14)),
            "freteDevido": num(c(15)),
            "margem": num(c(16)),
            "margemUnitaria": num(c(17)),
            "pendencias": num(c(18), 0),
            "status": c(19),
        })

    # ---------------- TRANSPORTADORAS + TARIFAS ----------------
    ws = wb["TRANSPORTADORAS"]
    transportadoras = []
    for r in linhas_ate_branco(ws, header_row=6, ultima_col=4):
        c = lambda col: txt(ws.cell(row=r, column=col).value)
        transportadoras.append({
            "nome": c(1),
            "cnpj": c(2),
            "prazoPagamento": c(3),
            "status": c(4),
        })

    tarifas_frete = []
    for r in linhas_ate_branco(ws, header_row=31, ultima_col=9):
        c = lambda col: txt(ws.cell(row=r, column=col).value)
        tarifas_frete.append({
            "transportadora": c(2),
            "local": c(3),
            "tarifaCheia": num(c(4)),
            "tarifaCte": num(c(5)),
            "tarifaMotorista": num(c(6)),
            "vigencia": data_iso(c(7)),
            "situacao": c(9),
        })

    # ---------------- MOTORISTAS + PREÇOS DE COMBUSTÍVEL ----------------
    ws = wb["MOTORISTAS"]
    motoristas = []
    for r in linhas_ate_branco(ws, header_row=6, ultima_col=23):
        c = lambda col: txt(ws.cell(row=r, column=col).value)
        motoristas.append({
            "nome": c(1),
            "transportadora": c(2),
            "placaCavalo": c(3),
            "placaCarreta": c(4),
            "cubagem": num(c(5), 0),
            "status": c(6),
            "entregas": num(c(9), 0),
            "credito": num(c(10)),
            "adiantamentos": num(c(11)),
            "combustivel": num(c(13)),
            "liquido": num(c(14)),
            "valorPago": num(c(15)),
            "saldoAPagar": num(c(16)),
            "debitoNaoCompensado": num(c(17)),
            "situacaoFatura": c(20),
            "saldoAberto": num(c(23)),
        })

    precos_combustivel = []
    for r in linhas_ate_branco(ws, header_row=71, ultima_col=5):
        c = lambda col: txt(ws.cell(row=r, column=col).value)
        precos_combustivel.append({
            "vigencia": data_iso(c(1)),
            "precoLitro": num(c(2)),
            "combustivel": c(3),
            "situacao": c(5),
        })

    # ---------------- AUDITORIA ----------------
    # A própria aba já traz, nas linhas 211-231, um resumo pré-calculado por
    # fórmulas ("RESUMO DE OCORRÊNCIAS POR TIPO DE VERIFICAÇÃO") — é mais
    # confiável usar esse resumo do que recontar as 18 colunas de checagem
    # por embarque (que ficam nas linhas 7-~54, uma faixa diferente da aba).
    ws = wb["AUDITORIA"]
    auditoria_resumo = []
    for r in linhas(ws, header_row=211):
        c = lambda col: txt(ws.cell(row=r, column=col).value)
        auditoria_resumo.append({
            "verificacao": limpa_header(c(1)),
            "ocorrencias": num(c(2), 0),
            "situacao": c(3),
        })

    # ---------------- CLIENTES E LOCAIS ----------------
    ws = wb["CLIENTES E LOCAIS"]
    clientes_locais = []
    for r in linhas(ws, header_row=6):
        c = lambda col: txt(ws.cell(row=r, column=col).value)
        clientes_locais.append({
            "local": c(1),
            "cliente": c(2),
            "cidadeUf": c(3),
            "exigeAgendamento": c(7),
        })

    return {
        "meta": meta,
        "embarques": embarques,
        "contratos": contratos,
        "posicaoContrato": posicao_contrato,
        "transportadoras": transportadoras,
        "tarifasFrete": tarifas_frete,
        "motoristas": motoristas,
        "auditoriaResumo": auditoria_resumo,
        "clientesLocais": clientes_locais,
        "precosCombustivel": precos_combustivel,
    }


def injetar_no_html(dados: dict, origem: Path, destino: Path):
    html = origem.read_text(encoding="utf-8")
    novo_json = json.dumps(dados, ensure_ascii=False, separators=(",", ":"))
    html_novo, n = re.subn(
        r"const DATA = \{.*?\};",
        "const DATA = " + novo_json.replace("\\", "\\\\") + ";",
        html,
        count=1,
        flags=re.DOTALL,
    )
    if n != 1:
        raise SystemExit("Não encontrei a linha 'const DATA = {...};' em " + str(origem))
    destino.write_text(html_novo, encoding="utf-8")


def comparar_com_atual(dados: dict) -> bool:
    if not JSON_PATH.exists():
        print("Nenhum dashboard_slim.json anterior para comparar.")
        return True
    atual = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    igual = dados == atual
    if igual:
        print("OK: dados gerados são idênticos ao dashboard_slim.json atual.")
    else:
        print("ATENÇÃO: diferenças em relação ao dashboard_slim.json atual (normal se a planilha mudou).")
    return igual


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="só compara, não grava nada")
    args = parser.parse_args()

    xlsx_path = achar_planilha()
    print(f"Lendo planilha: {xlsx_path.name}")
    dados = gerar(xlsx_path)
    comparar_com_atual(dados)

    if args.check:
        return

    JSON_PATH.write_text(
        json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    injetar_no_html(dados, HTML_SOURCE, HTML_SOURCE)
    injetar_no_html(dados, HTML_SOURCE, HTML_PUBLISH)
    print(f"Gerado: {JSON_PATH.name}, {HTML_SOURCE.name} (atualizado) e {HTML_PUBLISH.name} (para publicar)")


if __name__ == "__main__":
    sys.exit(main())
