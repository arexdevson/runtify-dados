#!/usr/bin/env python3
"""
Runtify · resumo de segurança da cidade de São Paulo (dados públicos da SSP-SP).

Baixa a planilha "SP Dados Criminais" do ano (e do anterior, se faltar mês), fica só com
ROUBO e FURTO em VIA PÚBLICA na capital, dos últimos 3 meses publicados, e conta as
ocorrências em quadradinhos de ~500 m. O app baixa este resumo da cidade inteira (nunca
pergunta por um endereço), então nada revela onde o usuário está.

Uso:  python gerar_mapa.py [--planilha arquivo.xlsx ...] [--saida sp-capital.json]
Precisa de: openpyxl
"""
import argparse
import datetime as dt
import json
import sys
import tempfile
import time
import urllib.request
from collections import defaultdict
from pathlib import Path

import openpyxl

INDICE = 'https://www.ssp.sp.gov.br/assets/estatistica/transparencia/spDados160_516.json'
BASE = 'https://www.ssp.sp.gov.br/'
UA = {'User-Agent': 'Runtify-dados/1.0 (resumo publico de seguranca; github)'}

PASSO = 0.005          # tamanho do quadradinho em graus (~550 m x ~510 m em SP)
MESES_JANELA = 3       # últimos meses publicados que entram no resumo
MUNICIPIO = 'S.PAULO'
CRIMES = {'Roubo (art. 157)': 'roubo', 'Furto (art. 155)': 'furto'}
# lat/lon plausíveis para a capital (descarta coordenadas zeradas ou erradas)
CAIXA = (-24.01, -23.35, -46.83, -46.36)


def com_tentativas(funcao, tentativas=4):
    """O site da SSP às vezes derruba a conexão: tenta de novo, esperando cada vez mais."""
    for n in range(tentativas):
        try:
            return funcao()
        except OSError as erro:
            if n == tentativas - 1:
                raise
            espera = 60 * (n + 1)
            print(f'falhou ({erro}), tentando de novo em {espera}s', file=sys.stderr)
            time.sleep(espera)


def baixar(url, destino):
    def uma_vez():
        req = urllib.request.Request(url, headers=UA)
        with urllib.request.urlopen(req, timeout=600) as r, open(destino, 'wb') as f:
            while bloco := r.read(1 << 20):
                f.write(bloco)
    com_tentativas(uma_vez)


def arquivos_dos_anos(anos):
    def uma_vez():
        req = urllib.request.Request(INDICE, headers=UA)
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode('utf-8-sig'))
    indice = com_tentativas(uma_vez)
    # alguns itens do índice não são um ano (ex.: "Dados a partir de 2013"): ficam de fora
    por_ano = {int(x['periodo']): BASE + x['arquivo'] for g in indice['data'] for x in g['lista']
               if str(x.get('periodo', '')).strip().isdigit()}
    return [por_ano[a] for a in anos if a in por_ano]


def periodo_do_dia(hora):
    """0 madrugada (0-5h), 1 manhã (6-11h), 2 tarde (12-17h), 3 noite (18-23h)."""
    if hora is None or not hasattr(hora, 'hour'):
        return None
    return hora.hour // 6


def numero(v):
    try:
        return float(str(v).replace(',', '.'))
    except (TypeError, ValueError):
        return 0.0


def ler_ocorrencias(planilha):
    """Gera (ano, mes, tipo, lat, lon, periodo) das linhas que interessam."""
    wb = openpyxl.load_workbook(planilha, read_only=True)
    for aba in wb.sheetnames:
        linhas = wb[aba].iter_rows(values_only=True)
        cab = next(linhas, None)
        if not cab or 'RUBRICA' not in cab:
            continue  # aba de explicação dos campos
        ix = {c: i for i, c in enumerate(cab) if c}
        for r in linhas:
            if r[ix['NOME_MUNICIPIO']] != MUNICIPIO or r[ix['DESCR_TIPOLOCAL']] != 'Via Pública':
                continue
            tipo = CRIMES.get(str(r[ix['RUBRICA']]).strip())
            if not tipo:
                continue
            lat, lon = numero(r[ix['LATITUDE']]), numero(r[ix['LONGITUDE']])
            if not (CAIXA[0] <= lat <= CAIXA[1] and CAIXA[2] <= lon <= CAIXA[3]):
                continue
            yield (int(r[ix['ANO_ESTATISTICA']]), int(r[ix['MES_ESTATISTICA']]), tipo, lat, lon,
                   periodo_do_dia(r[ix['HORA_OCORRENCIA_BO']]))


def percentil(valores, p):
    v = sorted(valores)
    return v[min(len(v) - 1, int(len(v) * p))] if v else 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--planilha', nargs='*', help='usar planilhas já baixadas em vez de baixar')
    ap.add_argument('--saida', default='sp-capital.json')
    args = ap.parse_args()

    if args.planilha:
        planilhas = [Path(p) for p in args.planilha]
    else:
        ano = dt.date.today().year
        pasta = Path(tempfile.mkdtemp())
        planilhas = []
        for i, url in enumerate(arquivos_dos_anos([ano - 1, ano])):
            destino = pasta / f'ssp_{i}.xlsx'
            print('baixando', url, file=sys.stderr)
            baixar(url, destino)
            planilhas.append(destino)

    ocorrencias = [o for p in planilhas for o in ler_ocorrencias(p)]
    meses = sorted({(o[0], o[1]) for o in ocorrencias})[-MESES_JANELA:]
    janela = set(meses)

    # quadradinho -> [roubos, furtos, madrugada, manhã, tarde, noite]
    celulas = defaultdict(lambda: [0, 0, 0, 0, 0, 0])
    for ano, mes, tipo, lat, lon, per in ocorrencias:
        if (ano, mes) not in janela:
            continue
        c = celulas[f'{round(lat / PASSO)},{round(lon / PASSO)}']
        c[0 if tipo == 'roubo' else 1] += 1
        if per is not None:
            c[2 + per] += 1

    # "peso" da vizinhança (3x3 quadradinhos ≈ 1,5 km): roubo vale 2, furto vale 1
    def peso(chave):
        i, j = map(int, chave.split(','))
        total = 0
        for di in (-1, 0, 1):
            for dj in (-1, 0, 1):
                c = celulas.get(f'{i + di},{j + dj}')
                if c:
                    total += 2 * c[0] + c[1]
        return total

    pesos = [peso(k) for k in celulas]
    resumo = {
        'fonte': 'SSP-SP, SP Dados Criminais (dados públicos)',
        'gerado_em': dt.date.today().isoformat(),
        'meses': [f'{a}-{m:02d}' for a, m in meses],
        'passo_graus': PASSO,
        'crimes': 'roubo e furto em via pública',
        # limites para os selos, comparando com a própria cidade
        'limites': {'atencao': percentil(pesos, 0.60), 'redobrada': percentil(pesos, 0.90)},
        'celulas': dict(celulas),
    }
    Path(args.saida).write_text(json.dumps(resumo, separators=(',', ':')), encoding='utf-8')
    print(f"ok: {len(celulas)} quadradinhos, meses {resumo['meses']}, limites {resumo['limites']}", file=sys.stderr)


if __name__ == '__main__':
    main()
