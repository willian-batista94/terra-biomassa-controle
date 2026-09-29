"""Gera dashboard_slim.json e injeta os dados no dashboard.html / index.html a
partir da planilha de controle (dados/*.xlsx).

Uso:
    python gerar_dashboard.py                  # gera, cifra e já publica (git add + commit + push)
    python gerar_dashboard.py --check          # só valida, sem gravar nada
    python gerar_dashboard.py --sem-publicar   # gera e cifra, mas não publica

A senha de abertura do painel é lida do arquivo local senha_painel.txt (fora
do git) ou da variável TB_SENHA; só é perguntada no terminal se nenhum dos
dois existir. O index.html publicado leva os dados cifrados (PBKDF2-SHA256 +
AES-256-GCM) e só o navegador, com a senha, consegue abri-los. Para trocar a
senha, edite senha_painel.txt e rode o script de novo. O dashboard.html (local, fora do git) mantém os
dados abertos para desenvolvimento.

Robustez contra novas versões da planilha:
  * cada coluna lida tem o cabeçalho esperado declarado nas SPEC_*; se a
    planilha mudar de layout, validar_cabecalhos() aborta com a lista exata
    do que mudou, em vez de gerar números errados em silêncio;
  * tabelas que ficam no meio de uma aba (tarifas, combustível, resumo da
    auditoria, títulos a receber) são localizadas pelo texto do cabeçalho,
    não por número de linha;
  * conferir_painel() recalcula os indicadores com as mesmas regras do
    dashboard e compara com a aba PAINEL da própria planilha.
"""

import argparse
import base64
import getpass
import json
import os
import re
import subprocess
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

import openpyxl

BASE_DIR = Path(__file__).resolve().parent
DADOS_DIR = BASE_DIR / "dados"
JSON_PATH = BASE_DIR / "dashboard_slim.json"
HIST_PATH = BASE_DIR / "historico_kpis.json"
HTML_SOURCE = BASE_DIR / "dashboard.html"
HTML_PUBLISH = BASE_DIR / "index.html"
SENHA_PATH = BASE_DIR / "senha_painel.txt"

PBKDF2_ITER = 600_000


# ---------------------------------------------------------------- helpers

def achar_planilha() -> Path:
    arquivos = sorted(DADOS_DIR.glob("*.xlsx"))
    if not arquivos:
        raise SystemExit(f"Nenhuma planilha .xlsx encontrada em {DADOS_DIR}")
    return arquivos[-1]


def num(v, casas=2):
    if v is None or v == "" or isinstance(v, str):
        return None if v == "" else v
    r = round(float(v), casas)
    return int(r) if r == int(r) else r


def txt(v):
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def data_iso(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str) and re.fullmatch(r"\d{2}/\d{2}/\d{4}", v.strip()):
        return datetime.strptime(v.strip(), "%d/%m/%Y").strftime("%Y-%m-%d")
    return txt(v)


def limpa_header(texto):
    return re.sub(r"\s+", " ", str(texto)).strip()


def norm(texto):
    """Comparação tolerante de cabeçalhos: sem acento, maiúsculo, espaços simples."""
    s = unicodedata.normalize("NFKD", limpa_header(texto or ""))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.upper()


def vazio(v):
    return v is None or (isinstance(v, str) and not v.strip())


def linhas(ws, header_row, key_col=1):
    """Uma linha por registro: pula linhas com a coluna-chave vazia, até o fim da aba."""
    for r in range(header_row + 1, ws.max_row + 1):
        if not vazio(ws.cell(row=r, column=key_col).value):
            yield r


def linhas_ate_branco(ws, header_row, ultima_col):
    """Para abas com duas tabelas: para na primeira linha totalmente vazia."""
    for r in range(header_row + 1, ws.max_row + 1):
        if all(vazio(ws.cell(row=r, column=c).value) for c in range(1, ultima_col + 1)):
            return
        yield r


def achar_linha(ws, texto, col=1, depois_de=0):
    alvo = norm(texto)
    for r in range(depois_de + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, str) and norm(v).startswith(alvo):
            return r
    raise SystemExit(f"Não encontrei '{texto}' na coluna {col} da aba {ws.title}.")


CONVERSORES = {
    "t": txt,
    "d": data_iso,
    "n": lambda v: num(v, 2),
    "n0": lambda v: num(v, 0),
    "n3": lambda v: num(v, 3),
    "n4": lambda v: num(v, 4),
}


def ler_registro(ws, r, spec):
    return {chave: CONVERSORES[tipo](ws.cell(row=r, column=col).value)
            for chave, col, _cab, tipo in spec}


# ---------------------------------------------------------------- mapas de colunas
# (chave no JSON, coluna 1-indexada, cabeçalho esperado, tipo)

SPEC_EMBARQUES = [
    ("num", 1, "EMBARQUE Nº", "n0"),
    ("contrato", 2, "CONTRATO", "t"),
    ("cliente", 3, "CLIENTE", "t"),
    ("local", 4, "LOCAL DE DESCARGA", "t"),
    ("precoContratado", 6, "PREÇO CONTRATADO (R$/M³)", "n"),
    ("dataProgramada", 7, "DATA PROGRAMADA", "d"),
    ("volProg", 8, "VOLUME PROGRAMADO (M³)", "n"),
    ("transportadora", 9, "TRANSPORTADORA", "t"),
    ("placaCavalo", 10, "PLACA CAVALO", "t"),
    ("motorista", 12, "MOTORISTA", "t"),
    ("statusProgramacao", 13, "STATUS PROGRAMAÇÃO", "t"),
    ("dataCarreg", 14, "DATA CARREGAMENTO", "d"),
    ("volEmb", 16, "VOLUME EMBARCADO (M³)", "n"),
    ("numNfe", 21, "Nº NF-E", "t"),
    ("volNf", 24, "VOLUME NF (M³)", "n"),
    ("valorNf", 26, "VALOR NF (R$)", "n"),
    ("statusDocumentos", 30, "STATUS DOCUMENTOS", "t"),
    ("statusAgendamento", 34, "STATUS AGENDAMENTO", "t"),
    ("valorCte", 37, "VALOR CT-E (R$)", "n"),
    ("statusCteMdfe", 41, "STATUS CT-E / MDF-E", "t"),
    ("dataDescarga", 42, "DATA DESCARGA", "d"),
    ("volDescarregado", 44, "VOLUME DESCARREGADO (M³)", "n3"),
    ("volLiquido", 47, "VOLUME LÍQUIDO ACEITO (M³)", "n"),
    ("difCubM3", 48, "DIFERENÇA DE CUBAGEM (M³)", "n"),
    ("difCubPct", 49, "DIFERENÇA DE CUBAGEM (%)", "n4"),
    ("statusDescarga", 51, "STATUS DESCARGA", "t"),
    ("nfComplementar", 53, "Nº NF COMPLEMENTAR", "t"),
    ("volComplemento", 55, "VOLUME COMPLEMENTO (M³)", "n"),
    ("valorComplemento", 57, "VALOR COMPLEMENTO (R$)", "n"),
    ("statusComplemento", 59, "STATUS COMPLEMENTO", "t"),
    ("volFaturado", 60, "VOLUME TOTAL FATURADO (M³)", "n"),
    ("difAFaturar", 61, "DIFERENÇA A FATURAR (M³)", "n"),
    ("valorAFaturar", 62, "VALOR A FATURAR (R$)", "n"),
    ("receitaTotal", 63, "RECEITA TOTAL (R$)", "n"),
    ("statusConciliacao", 64, "STATUS CONCILIAÇÃO", "t"),
    ("tarifaCte", 65, "TARIFA TRANSPORTADORA CT-E (R$/M³)", "n"),
    ("tarifaCheia", 66, "TARIFA CHEIA (R$/M³)", "n"),
    ("baseFrete", 68, "BASE DE CÁLCULO DO FRETE (M³)", "n"),
    ("freteApurado", 69, "FRETE APURADO TOTAL (R$)", "n"),
    ("aPagarTransp", 70, "A PAGAR À TRANSPORTADORA CT-E (R$)", "n"),
    ("dataPgtoTransp", 71, "DATA PAGAMENTO TRANSPORTADORA", "d"),
    ("valorPagoTransp", 72, "VALOR PAGO TRANSPORTADORA (R$)", "n"),
    ("statusPgtoTransp", 73, "STATUS PGTO TRANSPORTADORA", "t"),
    ("creditoMotorista", 75, "CRÉDITO DO MOTORISTA FRETE SEM CT-E (R$)", "n"),
    ("adiantamento", 76, "ADIANTAMENTO (R$)", "n"),
    ("litros", 78, "QUANTIDADE LITROS", "n"),
    ("valorCombustivel", 81, "VALOR COMBUSTÍVEL (R$)", "n"),
    ("combustivelAvulso", 82, "COMBUSTÍVEL AVULSO VINCULADO (R$)", "n"),
    ("liquidoCarga", 83, "LÍQUIDO DA CARGA (R$)", "n"),
    ("saldoMotorista", 84, "SALDO A PAGAR MOTORISTA (R$)", "n"),
    ("dataPgtoMotorista", 85, "DATA PAGAMENTO MOTORISTA", "d"),
    ("valorPagoMotorista", 86, "VALOR PAGO MOTORISTA (R$)", "n"),
    ("statusPgtoMotorista", 87, "STATUS PGTO MOTORISTA", "t"),
    ("statusCiclo", 88, "STATUS DO CICLO", "t"),
    ("ocorrencias", 89, "PENDÊNCIAS", "n0"),
    ("dataBase", 91, "DATA-BASE (PERÍODO DO PAINEL)", "d"),
    ("dataCompetencia", 98, "DATA DE COMPETÊNCIA FINANCEIRA", "d"),
    ("vencimentoFrete", 99, "VENCIMENTO DO FRETE (TRANSPORTADORA)", "d"),
    ("dataRecebimento", 100, "DATA DO RECEBIMENTO (LOTE)", "d"),
    ("valorRecebido", 101, "VALOR RECEBIDO NO LOTE (R$)", "n"),
    ("saldoReceber", 102, "SALDO A RECEBER (R$)", "n"),
    ("statusRecebimento", 103, "STATUS DO RECEBIMENTO", "t"),
]

SPEC_CONTRATOS = [
    ("contrato", 1, "Nº CONTRATO", "t"),
    ("cliente", 2, "CLIENTE", "t"),
    ("produto", 4, "PRODUTO", "t"),
    ("volumeContratado", 5, "VOLUME CONTRATADO (M³)", "n"),
    ("preco", 6, "PREÇO (R$/M³)", "n"),
    ("inicioVigencia", 7, "INÍCIO DA VIGÊNCIA", "d"),
    ("fimVigencia", 8, "FIM DA VIGÊNCIA", "d"),
    ("localPrincipal", 9, "LOCAL PRINCIPAL DE DESCARGA", "t"),
    ("toleranciaPct", 10, "TOLERÂNCIA DE VOLUME (%)", "n"),
    ("condicaoPagamento", 11, "CONDIÇÃO DE PAGAMENTO", "t"),
    ("status", 13, "STATUS", "t"),
    ("prazoRecebimento", 16, "PRAZO DE RECEBIMENTO (DIAS)", "n0"),
    ("baseVencimento", 17, "BASE DO VENCIMENTO", "t"),
]

SPEC_POSICAO = [
    ("contrato", 1, "CONTRATO", "t"),
    ("cliente", 2, "CLIENTE", "t"),
    ("volumeContratado", 4, "VOLUME CONTRATADO (M³)", "n"),
    ("volumeLiquido", 10, "VOLUME LÍQUIDO ACEITO (M³)", "n"),
    ("saldoEntregar", 11, "SALDO A ENTREGAR (M³)", "n"),
    ("pctExecutado", 12, "% EXECUTADO", "n4"),
    ("receitaFaturada", 13, "RECEITA FATURADA (R$)", "n"),
    ("valorAFaturar", 14, "VALOR A FATURAR (R$)", "n"),
    ("freteDevido", 15, "FRETE DEVIDO (R$)", "n"),
    ("margem", 16, "MARGEM DE CONTRIBUIÇÃO (R$)", "n"),
    ("margemUnitaria", 17, "MARGEM UNITÁRIA (R$/M³)", "n"),
    ("pendencias", 18, "PENDÊNCIAS", "n0"),
    ("status", 19, "STATUS DO CONTRATO", "t"),
    ("recebido", 20, "RECEBIDO (R$)", "n"),
    ("saldoReceber", 21, "SALDO A RECEBER (R$)", "n"),
]

SPEC_TRANSPORTADORAS = [
    ("nome", 1, "TRANSPORTADORA", "t"),
    ("cnpj", 2, "CNPJ", "t"),
    ("prazoPagamento", 3, "PRAZO DE PAGAMENTO", "t"),
    ("status", 4, "STATUS", "t"),
]

SPEC_TARIFAS = [
    ("transportadora", 2, "TRANSPORTADORA", "t"),
    ("local", 3, "LOCAL DE DESCARGA", "t"),
    ("tarifaCheia", 4, "TARIFA CHEIA (R$/M³)", "n"),
    ("tarifaCte", 5, "TARIFA TRANSPORTADORA CT-E (R$/M³)", "n"),
    ("tarifaMotorista", 6, "TARIFA MOTORISTA (R$/M³) · REFERÊNCIA", "n"),
    ("vigencia", 7, "VIGÊNCIA A PARTIR DE", "d"),
    ("situacao", 9, "SITUAÇÃO", "t"),
]

# Só o cadastro: os totais por motorista são calculados no dashboard a partir
# dos embarques, porque as colunas de totais da aba MOTORISTAS dependem do
# período que estiver digitado no PAINEL quando a planilha foi salva.
SPEC_MOTORISTAS = [
    ("nome", 1, "MOTORISTA", "t"),
    ("transportadora", 2, "TRANSPORTADORA", "t"),
    ("placaCavalo", 3, "PLACA CAVALO", "t"),
    ("placaCarreta", 4, "PLACA CARRETA", "t"),
    ("cubagem", 5, "CUBAGEM DO CONJUNTO (M³)", "n0"),
    ("status", 6, "STATUS", "t"),
]

SPEC_COMBUSTIVEL = [
    ("vigencia", 1, "VIGÊNCIA A PARTIR DE", "d"),
    ("precoLitro", 2, "PREÇO POR LITRO (R$/L)", "n"),
    ("combustivel", 3, "COMBUSTÍVEL", "t"),
    ("situacao", 5, "SITUAÇÃO", "t"),
]

SPEC_CLIENTES = [
    ("local", 1, "LOCAL DE DESCARGA", "t"),
    ("cliente", 2, "CLIENTE", "t"),
    ("cidadeUf", 3, "CIDADE / UF", "t"),
    ("distanciaKm", 4, "DISTÂNCIA (KM)", "n0"),
    ("janela", 6, "JANELA DE RECEBIMENTO", "t"),
    ("exigeAgendamento", 7, "EXIGE AGENDAMENTO", "t"),
]

SPEC_TITULOS = [
    ("num", 1, "EMBARQUE Nº", "n0"),
    ("contrato", 2, "CONTRATO", "t"),
    ("cliente", 3, "CLIENTE", "t"),
    ("dataCompetencia", 4, "DATA DE COMPETÊNCIA", "d"),
    ("prazo", 5, "PRAZO (DIAS)", "n0"),
    ("numNfe", 7, "Nº NF-E", "t"),
    ("valorNfe", 8, "VALOR NF-E (R$)", "n"),
    ("vencimentoNfe", 9, "VENCIMENTO NF-E", "d"),
    ("recebidoNfe", 10, "RECEBIDO NF-E (R$)", "n"),
    ("saldoNfe", 11, "SALDO NF-E (R$)", "n"),
    ("numNfCompl", 12, "Nº NF COMPLEMENTAR", "t"),
    ("valorNfCompl", 13, "VALOR NF COMPL. (R$)", "n"),
    ("vencimentoNfCompl", 14, "VENCIMENTO NF COMPL.", "d"),
    ("recebidoNfCompl", 15, "RECEBIDO NF COMPL. (R$)", "n"),
    ("saldoNfCompl", 16, "SALDO NF COMPL. (R$)", "n"),
    ("faturado", 17, "TOTAL FATURADO (R$)", "n"),
    ("recebido", 18, "TOTAL RECEBIDO (R$)", "n"),
    ("saldo", 19, "SALDO A RECEBER (R$)", "n"),
    ("vencido", 20, "VENCIDO (R$)", "n"),
    ("aVencer", 21, "A VENCER (R$)", "n"),
    ("aFaturarSemNf", 22, "A FATURAR SEM NF (R$)", "n"),
    ("diasAtraso", 23, "DIAS EM ATRASO", "n0"),
    ("status", 24, "STATUS DO RECEBIMENTO", "t"),
    ("alerta", 25, "ALERTA DA NOTA", "t"),
]

SPEC_RECEBIMENTOS = [
    ("data", 1, "DATA DO RECEBIMENTO", "d"),
    ("competenciaDe", 3, "COMPETÊNCIA DE", "d"),
    ("competenciaAte", 4, "COMPETÊNCIA ATÉ", "d"),
    ("nf", 5, "Nº NF (NOTA ESPECÍFICA)", "t"),
    ("valor", 6, "VALOR RECEBIDO (R$)", "n"),
    ("forma", 7, "FORMA DE RECEBIMENTO", "t"),
    ("tipo", 10, "TIPO DO LANÇAMENTO", "t"),
    ("contrato", 11, "CONTRATO", "t"),
    ("cliente", 12, "CLIENTE", "t"),
    ("embarques", 13, "EMBARQUES VINCULADOS", "n0"),
    ("valorTitulos", 14, "VALOR DOS TÍTULOS (R$)", "n"),
    ("diferenca", 16, "DIFERENÇA DO LOTE (R$)", "n"),
    ("situacao", 17, "SITUAÇÃO DO LANÇAMENTO", "t"),
]

SPEC_NOTAS_COMPL = [
    ("nota", 1, "Nº DA NOTA COMPLEMENTAR", "t"),
    ("dataEmissao", 2, "DATA DE EMISSÃO", "d"),
    ("contrato", 4, "CONTRATO", "t"),
    ("descargasDe", 5, "DESCARGAS DE", "d"),
    ("descargasAte", 6, "DESCARGAS ATÉ", "d"),
    ("motivo", 7, "MOTIVO", "t"),
    ("volumeNota", 8, "VOLUME DA NOTA (M³)", "n"),
    ("embarques", 9, "EMBARQUES VINCULADOS", "n0"),
    ("volumeRateado", 10, "VOLUME RATEADO (M³)", "n"),
    ("diferenca", 11, "DIFERENÇA (M³)", "n"),
    ("valorRateado", 12, "VALOR RATEADO (R$)", "n"),
    ("situacao", 13, "SITUAÇÃO", "t"),
]

SPEC_PGTO_TRANSP = [
    ("data", 1, "DATA DO PAGAMENTO", "d"),
    ("transportadora", 2, "TRANSPORTADORA", "t"),
    ("de", 3, "COMPETÊNCIA DE", "d"),
    ("ate", 4, "COMPETÊNCIA ATÉ", "d"),
    ("valor", 5, "VALOR PAGO (R$)", "n"),
    ("forma", 6, "FORMA DE PAGAMENTO", "t"),
    ("cteVinculados", 7, "CT-E VINCULADOS", "n0"),
    ("valorCte", 8, "VALOR DOS CT-E (R$)", "n"),
    ("diferenca", 9, "DIFERENÇA (R$)", "n"),
    ("situacao", 10, "SITUAÇÃO", "t"),
]

SPEC_PGTO_MOT = [
    ("data", 1, "DATA DO PAGAMENTO", "d"),
    ("motorista", 2, "MOTORISTA", "t"),
    ("de", 3, "PERÍODO DE", "d"),
    ("ate", 4, "PERÍODO ATÉ", "d"),
    ("valor", 5, "VALOR PAGO (R$)", "n"),
    ("forma", 6, "FORMA DE PAGAMENTO", "t"),
    ("cargas", 7, "CARGAS VINCULADAS", "n0"),
    ("liquidoCargas", 8, "LÍQUIDO DAS CARGAS (R$)", "n"),
    ("diferenca", 9, "DIFERENÇA (R$)", "n"),
    ("situacao", 10, "SITUAÇÃO", "t"),
]

SPEC_AVULSOS = [
    ("controle", 1, "Nº CONTROLE ABASTECIMENTO", "t"),
    ("motorista", 2, "MOTORISTA", "t"),
    ("data", 3, "DATA ABASTECIMENTO", "d"),
    ("litros", 4, "QUANTIDADE LITROS", "n"),
    ("transportadora", 7, "TRANSPORTADORA", "t"),
    ("precoLitro", 8, "PREÇO/LT (R$)", "n"),
    ("valor", 9, "VALOR (R$)", "n"),
    ("descontadoEmbarque", 10, "DESCONTADO NO EMBARQUE Nº", "n0"),
    ("dataCompetencia", 12, "DATA DE COMPETÊNCIA DO DESCONTO", "d"),
    ("status", 13, "STATUS", "t"),
]

SPEC_AUDITORIA = [
    ("verificacao", 1, "VERIFICAÇÃO", "t"),
    ("ocorrencias", 2, "OCORRÊNCIAS", "n0"),
    ("situacao", 3, "SITUAÇÃO", "t"),
]


def validar_cabecalhos(ws, header_row, spec, erros):
    for chave, col, esperado, _tipo in spec:
        achado = ws.cell(row=header_row, column=col).value
        if norm(achado) != norm(esperado):
            erros.append(
                f"  {ws.title} linha {header_row}, coluna {col} ({chave}): "
                f"esperado '{esperado}', encontrado '{limpa_header(achado) if achado else '(vazio)'}'"
            )


def tabela(ws, header_row, spec, erros, key_col=1, ate_branco=None):
    validar_cabecalhos(ws, header_row, spec, erros)
    it = (linhas_ate_branco(ws, header_row, ate_branco) if ate_branco
          else linhas(ws, header_row, key_col))
    return [ler_registro(ws, r, spec) for r in it]


# ---------------------------------------------------------------- extração

def gerar(xlsx_path: Path):
    """Devolve (dados, painel) — painel são os valores da aba PAINEL, usados só
    na conferência."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    erros = []

    # META
    capa = wb["CAPA"]
    versao = None
    for r in range(1, capa.max_row + 1):
        if norm(capa.cell(r, 2).value) == "VERSAO":
            versao = txt(capa.cell(r, 3).value)
    painel_ws = wb["PAINEL"]
    params_ws = wb["PARÂMETROS"]
    params = {}
    for r in range(1, params_ws.max_row + 1):
        rot = params_ws.cell(r, 1).value
        if isinstance(rot, str):
            params[norm(rot)] = params_ws.cell(r, 2).value
    meta = {
        "empresa": "TERRA BIOMASSA INDUSTRIA E COMERCIO LTDA",
        "produto": "BIOMASSA - CAVACO DE MADEIRA",
        "emissao": data_iso(capa["C24"].value),
        "versaoPlanilha": (versao or "").split("·")[0].strip() or None,
        "arquivo": xlsx_path.name,
        "geradoEm": datetime.now().strftime("%Y-%m-%dT%H:%M"),
        "parametros": {
            "tolCubagem": num(params.get("TOLERANCIA DE DIFERENCA DE CUBAGEM"), 4),
            "baseFrete": txt(params.get("BASE DE CALCULO DO FRETE")),
            "tolRecebimento": num(params.get("TOLERANCIA DE RECEBIMENTO (R$)")),
        },
    }
    if norm(capa["B24"].value) != "DATA DE EMISSAO":
        erros.append(f"  CAPA B24: esperado 'DATA DE EMISSÃO', encontrado '{capa['B24'].value}'")

    embarques = tabela(wb["EMBARQUES"], 6, SPEC_EMBARQUES, erros)
    contratos = tabela(wb["CONTRATOS"], 6, SPEC_CONTRATOS, erros)
    posicao = tabela(wb["POSIÇÃO POR CONTRATO"], 6, SPEC_POSICAO, erros)

    ws = wb["TRANSPORTADORAS"]
    transportadoras = tabela(ws, 6, SPEC_TRANSPORTADORAS, erros, ate_branco=4)
    h_tar = achar_linha(ws, "CHAVE (TRANSPORTADORA | LOCAL)")
    tarifas = tabela(ws, h_tar, SPEC_TARIFAS, erros, ate_branco=9)

    ws = wb["MOTORISTAS"]
    motoristas = tabela(ws, 6, SPEC_MOTORISTAS, erros, ate_branco=6)
    h_comb = achar_linha(ws, "VIGÊNCIA A PARTIR DE", depois_de=7)
    combustivel = tabela(ws, h_comb, SPEC_COMBUSTIVEL, erros, ate_branco=5)

    clientes = tabela(wb["CLIENTES E LOCAIS"], 6, SPEC_CLIENTES, erros)

    ws = wb["CONTAS A RECEBER"]
    h_tit = achar_linha(ws, "EMBARQUE Nº", depois_de=10)
    titulos = tabela(ws, h_tit, SPEC_TITULOS, erros)

    recebimentos = tabela(wb["RECEBIMENTOS"], 6, SPEC_RECEBIMENTOS, erros)
    notas = tabela(wb["NOTAS COMPLEMENTARES"], 6, SPEC_NOTAS_COMPL, erros)
    pgto_transp = tabela(wb["PAGAMENTOS TRANSPORTADORA"], 6, SPEC_PGTO_TRANSP, erros)
    pgto_mot = tabela(wb["PAGAMENTOS MOTORISTA"], 6, SPEC_PGTO_MOT, erros)
    avulsos = tabela(wb["ABASTECIMENTOS AVULSOS"], 6, SPEC_AVULSOS, erros, key_col=2)

    # AUDITORIA: o resumo pré-calculado fica abaixo das linhas por embarque;
    # lê do cabeçalho "VERIFICAÇÃO" até "TOTAL DE OCORRÊNCIAS", inclusive.
    ws = wb["AUDITORIA"]
    h_res = achar_linha(ws, "RESUMO DE OCORRÊNCIAS")
    h_aud = achar_linha(ws, "VERIFICAÇÃO", depois_de=h_res)
    validar_cabecalhos(ws, h_aud, SPEC_AUDITORIA, erros)
    auditoria = []
    for r in linhas(ws, h_aud):
        reg = ler_registro(ws, r, SPEC_AUDITORIA)
        reg["verificacao"] = limpa_header(reg["verificacao"])
        auditoria.append(reg)
        if norm(reg["verificacao"]) == "TOTAL DE OCORRENCIAS":
            break

    if erros:
        raise SystemExit(
            "A estrutura da planilha mudou; nada foi gerado. Ajuste as SPEC_* em "
            "gerar_dashboard.py:\n" + "\n".join(erros)
        )

    # Valores do PAINEL, por rótulo, para a conferência.
    painel = {
        "de": data_iso(painel_ws["B7"].value),
        "ate": data_iso(painel_ws["B8"].value),
        "valores": {},
    }
    for r in range(12, painel_ws.max_row + 1):
        rot, val = painel_ws.cell(r, 1).value, painel_ws.cell(r, 2).value
        if isinstance(rot, str) and isinstance(val, (int, float)):
            painel["valores"][limpa_header(rot)] = val

    dados = {
        "meta": meta,
        "embarques": embarques,
        "contratos": contratos,
        "posicaoContrato": posicao,
        "transportadoras": transportadoras,
        "tarifasFrete": tarifas,
        "motoristas": motoristas,
        "precosCombustivel": combustivel,
        "clientesLocais": clientes,
        "titulosReceber": titulos,
        "recebimentos": recebimentos,
        "notasComplementares": notas,
        "pagamentosTransportadora": pgto_transp,
        "pagamentosMotorista": pgto_mot,
        "abastecimentosAvulsos": avulsos,
        "auditoriaResumo": auditoria,
    }
    return dados, painel


# ---------------------------------------------------------------- indicadores
# Mesmas regras do dashboard (função indicadores() em dashboard.html) e da aba
# PAINEL: operacional pela DATA-BASE, financeiro pela DATA DE COMPETÊNCIA,
# recebimentos pela data do recebimento, "pago" pelas datas de pagamento.

def _no_periodo(d, de, ate):
    if de is None and ate is None:
        return True
    return d is not None and (de is None or d >= de) and (ate is None or d <= ate)


def indicadores(dados, de=None, ate=None):
    acumulado = de is None and ate is None
    E = dados["embarques"]
    s = lambda rows, k: sum((r.get(k) or 0) for r in rows if isinstance(r.get(k), (int, float)))
    op = [e for e in E if acumulado or _no_periodo(e["dataBase"], de, ate)]
    fin = [e for e in E if acumulado or _no_periodo(e["dataCompetencia"], de, ate)]
    tit = [t for t in dados["titulosReceber"] if acumulado or _no_periodo(t["dataCompetencia"], de, ate)]
    rec = [r for r in dados["recebimentos"] if acumulado or _no_periodo(r["data"], de, ate)]

    vol_emb = s(op, "volEmb")
    receita, frete = s(fin, "receitaTotal"), s(fin, "freteApurado")
    a_pagar_transp = s(fin, "aPagarTransp")
    saldo_transp = round(a_pagar_transp - s(fin, "valorPagoTransp"), 2)

    por_mot = {}
    for e in fin:
        if not e["motorista"]:
            continue
        liq = sum((e.get(k) or 0) * sinal for k, sinal in (
            ("creditoMotorista", 1), ("adiantamento", -1), ("valorCombustivel", -1),
            ("combustivelAvulso", -1), ("valorPagoMotorista", -1)))
        por_mot[e["motorista"]] = por_mot.get(e["motorista"], 0) + liq
    saldo_mot = round(sum(max(0, v) for v in por_mot.values()), 2)

    saldo_receber = s(tit, "saldo")
    pago = (sum(e["valorPagoTransp"] or 0 for e in E if _no_periodo(e["dataPgtoTransp"], de, ate) and e["dataPgtoTransp"])
            + sum(e["valorPagoMotorista"] or 0 for e in E if _no_periodo(e["dataPgtoMotorista"], de, ate) and e["dataPgtoMotorista"]))
    recebido = s(rec, "valor")
    vol_fat_fin = s(fin, "volFaturado")

    return {
        "embarques": len(op),
        "volProg": s(op, "volProg"),
        "volEmb": vol_emb,
        "volFaturado": s(op, "volFaturado"),
        "volDescarregado": s(op, "volDescarregado"),
        "volLiquido": s(op, "volLiquido"),
        "difCub": s(op, "difCubM3"),
        "emTransito": sum(1 for e in op if e["statusCiclo"] == "EM TRÂNSITO"),
        "ciclosConcluidos": sum(1 for e in op if e["statusCiclo"] == "CICLO CONCLUÍDO"),
        "ocorrencias": s(op, "ocorrencias"),
        "embComOcorrencia": sum(1 for e in op if (e["ocorrencias"] or 0) > 0),
        "conciliacoesAberto": sum(1 for e in op if e["statusConciliacao"] in ("FATURAR COMPLEMENTO", "SEM NF DE VENDA")),
        "faturadoMaior": sum(1 for e in op if e["statusConciliacao"] == "FATURADO A MAIOR"),
        "receita": receita,
        "aFaturar": s(fin, "valorAFaturar"),
        "precoMedio": receita / vol_fat_fin if vol_fat_fin else 0,
        "frete": frete,
        "aPagarTransp": a_pagar_transp,
        "saldoTransp": saldo_transp,
        "creditoMot": s(fin, "creditoMotorista"),
        "saldoMot": saldo_mot,
        "margem": round(receita - frete, 2),
        "faturado": s(tit, "faturado"),
        "recebido": recebido,
        "saldoReceber": saldo_receber,
        "vencido": s(tit, "vencido"),
        "aVencer": s(tit, "aVencer"),
        "aFaturarSemNf": s(tit, "aFaturarSemNf"),
        "posicaoLiquida": round(saldo_receber - saldo_transp - saldo_mot, 2),
        "pago": round(pago, 2),
        "caixa": round(recebido - pago, 2),
    }


CONFERENCIA = [
    # (indicador, rótulo na aba PAINEL)
    ("embarques", "EMBARQUES REGISTRADOS"),
    ("volProg", "VOLUME PROGRAMADO"),
    ("volEmb", "VOLUME EMBARCADO NA ORIGEM"),
    ("volFaturado", "VOLUME FATURADO — NF-E + COMPLEMENTO"),
    ("volDescarregado", "VOLUME DESCARREGADO NO DESTINO"),
    ("volLiquido", "VOLUME LÍQUIDO ACEITO NO DESTINO"),
    ("difCub", "DIFERENÇA DE CUBAGEM ACUMULADA"),
    ("emTransito", "EMBARQUES EM TRÂNSITO"),
    ("ciclosConcluidos", "CICLOS CONCLUÍDOS"),
    ("ocorrencias", "OCORRÊNCIAS DE AUDITORIA"),
    ("embComOcorrencia", "EMBARQUES COM OCORRÊNCIA"),
    ("conciliacoesAberto", "CONCILIAÇÕES EM ABERTO"),
    ("faturadoMaior", "EMBARQUES FATURADOS A MAIOR"),
    ("receita", "RECEITA FATURADA — NF-E + COMPLEMENTO"),
    ("aFaturar", "VALOR A FATURAR — COMPLEMENTO"),
    ("precoMedio", "PREÇO MÉDIO REALIZADO"),
    ("frete", "FRETE APURADO TOTAL"),
    ("aPagarTransp", "A PAGAR À TRANSPORTADORA — CT-E"),
    ("saldoTransp", "SALDO A PAGAR — TRANSPORTADORA"),
    ("creditoMot", "CRÉDITO AO MOTORISTA — FRETE SEM CT-E"),
    ("saldoMot", "SALDO A PAGAR — MOTORISTAS"),
    ("margem", "MARGEM DE CONTRIBUIÇÃO LOGÍSTICA"),
    ("faturado", "FATURADO — NF-E + COMPLEMENTO"),
    ("recebido", "RECEBIDO DE CLIENTES"),
    ("saldoReceber", "SALDO A RECEBER EM ABERTO"),
    ("vencido", "VENCIDO"),
    ("aVencer", "A VENCER"),
    ("aFaturarSemNf", "A FATURAR SEM NF — COMPLEMENTO PENDENTE"),
    ("posicaoLiquida", "(=) POSIÇÃO LÍQUIDA — A RECEBER MENOS A PAGAR"),
    ("pago", "(−) PAGO NO PERÍODO — TRANSPORTADORA E MOTORISTAS"),
    ("caixa", "(=) SALDO DE CAIXA DO PERÍODO"),
]


def conferir_painel(dados, painel) -> bool:
    calc = indicadores(dados, painel["de"], painel["ate"])
    vals = {norm(k): v for k, v in painel["valores"].items()}
    divergentes, ausentes = [], []
    for chave, rotulo in CONFERENCIA:
        esperado = vals.get(norm(rotulo))
        if esperado is None:
            ausentes.append(rotulo)
            continue
        if abs((calc[chave] or 0) - esperado) > 0.011:
            divergentes.append(f"  {rotulo}: painel {esperado:,.2f} × dashboard {calc[chave]:,.2f}")
    periodo = f"{painel['de'] or 'início'} a {painel['ate'] or 'fim'}"
    if divergentes:
        print(f"ATENÇÃO: conferência com o PAINEL ({periodo}) divergiu em {len(divergentes)} indicador(es):")
        print("\n".join(divergentes))
    else:
        print(f"OK: {len(CONFERENCIA) - len(ausentes)} indicadores batem com o PAINEL ({periodo}).")
    if ausentes:
        print("  (rótulos não encontrados no PAINEL: " + "; ".join(ausentes) + ")")
    return not divergentes


# ---------------------------------------------------------------- histórico

KPIS_HIST = ["embarques", "volLiquido", "receita", "frete", "margem", "saldoReceber",
             "saldoTransp", "saldoMot", "posicaoLiquida", "ocorrencias"]


def _snapshot_legado(dados):
    """Indicadores possíveis a partir de um dashboard_slim.json antigo (sem
    contas a receber nem data de competência)."""
    E = dados.get("embarques", [])
    s = lambda k: sum((e.get(k) or 0) for e in E if isinstance(e.get(k), (int, float)))
    receita, frete = s("receitaTotal"), s("freteApurado")
    return {"embarques": len(E), "volLiquido": s("volLiquido"), "receita": receita,
            "frete": frete, "margem": round(receita - frete, 2),
            "saldoTransp": round(s("aPagarTransp") - s("valorPagoTransp"), 2),
            "ocorrencias": s("ocorrencias")}


def atualizar_historico(dados, gravar=True):
    hist = {}
    if HIST_PATH.exists():
        hist = json.loads(HIST_PATH.read_text(encoding="utf-8"))
    else:
        # Primeira execução: reconstrói a partir das versões já publicadas.
        log = subprocess.run(["git", "log", "--format=%H", "--", "dashboard_slim.json"],
                             cwd=BASE_DIR, capture_output=True, text=True)
        for h in reversed(log.stdout.split()):
            show = subprocess.run(["git", "show", f"{h}:dashboard_slim.json"],
                                  cwd=BASE_DIR, capture_output=True, text=True, encoding="utf-8")
            if show.returncode != 0:
                continue
            try:
                antigo = json.loads(show.stdout)
            except json.JSONDecodeError:
                continue
            em = antigo.get("meta", {}).get("emissao")
            if em:
                hist[em] = _snapshot_legado(antigo)
    atual = indicadores(dados)
    hist[dados["meta"]["emissao"]] = {k: round(atual[k], 2) for k in KPIS_HIST}
    hist = dict(sorted(hist.items()))
    if gravar:
        HIST_PATH.write_text(json.dumps(hist, ensure_ascii=False, indent=2), encoding="utf-8")
    return hist


# ---------------------------------------------------------------- HTML / cifra

def _substituir_linha(html, prefixo, nova, origem):
    padrao = re.compile(r"^" + re.escape(prefixo) + r".*;$", re.MULTILINE)
    html_novo, n = padrao.subn(lambda _m: nova, html, count=1)
    if n != 1:
        raise SystemExit(f"Não encontrei a linha '{prefixo}...;' em {origem}")
    return html_novo


def cifrar(dados: dict, senha: str) -> dict:
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    salt, iv = os.urandom(16), os.urandom(12)
    chave = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt,
                       iterations=PBKDF2_ITER).derive(senha.encode("utf-8"))
    claro = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    ct = AESGCM(chave).encrypt(iv, claro, None)
    b64 = lambda b: base64.b64encode(b).decode("ascii")
    return {"v": 1, "iter": PBKDF2_ITER, "salt": b64(salt), "iv": b64(iv), "ct": b64(ct)}


def gravar_html_local(dados: dict):
    """dashboard.html: dados abertos (arquivo local, fora do git)."""
    html = HTML_SOURCE.read_text(encoding="utf-8")
    js = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = _substituir_linha(html, "const DATA = ", "const DATA = " + js + ";", HTML_SOURCE)
    html = _substituir_linha(html, "const PAYLOAD = ", "const PAYLOAD = null;", HTML_SOURCE)
    HTML_SOURCE.write_text(html, encoding="utf-8")


def gravar_html_publico(dados: dict, senha: str):
    """index.html: mesma página, dados só na forma cifrada."""
    html = HTML_SOURCE.read_text(encoding="utf-8")
    html = _substituir_linha(html, "const DATA = ", "const DATA = null;", HTML_SOURCE)
    payload = json.dumps(cifrar(dados, senha), separators=(",", ":"))
    html = _substituir_linha(html, "const PAYLOAD = ", "const PAYLOAD = " + payload + ";", HTML_SOURCE)
    # Trava de segurança: nenhum nome de cliente/motorista pode sair em texto.
    sensiveis = {c["cliente"] for c in dados["contratos"] if c["cliente"]}
    sensiveis |= {m["nome"] for m in dados["motoristas"] if m["nome"] and len(m["nome"]) > 4}
    vazados = sorted(n for n in sensiveis if n in html)
    if vazados:
        raise SystemExit("Abortado: o index.html conteria dados em texto aberto: " + ", ".join(vazados))
    HTML_PUBLISH.write_text(html, encoding="utf-8")


def obter_senha() -> str:
    """Senha de abertura do painel, usada para cifrar os dados. Ordem: variável
    TB_SENHA, arquivo local senha_painel.txt (fora do git) e, só se nenhum dos
    dois existir, pergunta uma vez no terminal e grava no arquivo."""
    senha = os.environ.get("TB_SENHA")
    if senha:
        return senha
    if SENHA_PATH.exists():
        senha = SENHA_PATH.read_text(encoding="utf-8").strip()
        if senha:
            return senha
    senha = getpass.getpass("Senha de abertura do painel (pedida só desta vez): ")
    if not senha:
        raise SystemExit("Senha vazia; nada foi publicado.")
    if getpass.getpass("Confirme a senha: ") != senha:
        raise SystemExit("As senhas não conferem; nada foi publicado.")
    SENHA_PATH.write_text(senha, encoding="utf-8")
    print(f"Senha gravada em {SENHA_PATH.name} (arquivo local, fora do git).")
    return senha


# ---------------------------------------------------------------- publicação

def publicar():
    """git add + commit + push só do index.html (o JSON aberto não vai ao git)."""
    def git(*args):
        return subprocess.run(["git", *args], cwd=BASE_DIR, capture_output=True, text=True)

    add = git("add", "index.html")
    if add.returncode != 0:
        raise SystemExit("Erro no 'git add':\n" + add.stderr)

    if git("diff", "--cached", "--quiet", "--", "index.html").returncode == 0:
        print("Nada para publicar: o index.html já bate com o último commit.")
        return

    mensagem = f"Atualiza dados — {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    commit = git("commit", "-m", mensagem, "--", "index.html")
    if commit.returncode != 0:
        raise SystemExit("Erro no 'git commit':\n" + commit.stderr)

    push = git("push")
    if push.returncode != 0:
        raise SystemExit("Commit criado localmente, mas o 'git push' falhou:\n" + push.stderr)

    print(f"Publicado: \"{mensagem}\"")
    print("O GitHub Pages atualiza em ~1 minuto.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="só valida e confere, não grava nada")
    parser.add_argument("--sem-publicar", action="store_true", dest="sem_publicar",
                        help="gera os arquivos mas não faz git add + commit + push")
    args = parser.parse_args()

    xlsx_path = achar_planilha()
    print(f"Lendo planilha: {xlsx_path.name}")
    dados, painel = gerar(xlsx_path)
    print(f"Estrutura OK: {len(dados['embarques'])} embarques, {len(dados['titulosReceber'])} títulos, "
          f"{len(dados['auditoriaResumo'])} linhas de auditoria.")
    conferir_painel(dados, painel)

    if args.check:
        return

    dados["historico"] = atualizar_historico(dados)
    JSON_PATH.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    gravar_html_local(dados)
    gravar_html_publico(dados, obter_senha())
    print(f"Gerado: {JSON_PATH.name}, {HTML_SOURCE.name} (dados abertos, local) e "
          f"{HTML_PUBLISH.name} (dados cifrados, para publicar)")

    if not args.sem_publicar:
        publicar()


if __name__ == "__main__":
    sys.exit(main())
