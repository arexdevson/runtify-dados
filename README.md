# runtify-dados

Dados **públicos** usados pelo app Runtify. Não há nenhum dado de usuário aqui.

## sp-capital.json

Resumo de **roubos e furtos em via pública** na cidade de São Paulo, nos últimos 3 meses publicados pela
Secretaria da Segurança Pública (SSP-SP, "SP Dados Criminais"). As ocorrências são contadas em quadradinhos
de ~500 m, separadas por período do dia (madrugada, manhã, tarde, noite).

O app baixa o arquivo da cidade inteira, então nunca informa onde o usuário está.

- Gerado por `gerar_mapa.py` (precisa de `openpyxl`).
- Atualizado automaticamente nos dias 10 e 25 de cada mês (`.github/workflows/atualizar.yml`).
- Fonte: SSP-SP, https://www.ssp.sp.gov.br/ (Estatística > Transparência)
