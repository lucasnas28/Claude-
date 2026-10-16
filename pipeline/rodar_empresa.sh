#!/bin/bash
# Monta o modelo de uma companhia americana do início ao fim.
# uso: pipeline/rodar_empresa.sh TICKER N_LINHAS_DRIVERS NOME_SAIDA
#   TICKER   = pasta em pipeline/empresas/ (ex.: orcl)
#   N        = linhas reservadas para o bloco de drivers no Model (ex.: 40; se o gen_drivers acusar falta, aumente)
#   NOME     = ex.: Oracle  -> grava Modelo_Oracle_2T26.xlsx na raiz do repositório
# variáveis: SEC_UA (obrigatória para baixar da SEC), PIPE_DATA (pasta de trabalho, padrão: dados/)
set -e
T=$1; N=$2; NOME=$3
P="$(cd "$(dirname "$0")" && pwd)"; REPO="$(cd "$P/.." && pwd)"
export PIPE_SEC="$P/sec"; export PIPE_DATA="$(mkdir -p "${PIPE_DATA:-$REPO/dados}" && cd "${PIPE_DATA:-$REPO/dados}" && pwd)"
MASK="$REPO/Modelo_2.0_proprio_Padrao_AJUSTADO-2028_corrigido.xlsx"
CFG="$P/empresas/$T/cfg.py"; SPEC="$P/empresas/$T/spec.py"
mkdir -p "$PIPE_DATA/$T"; cd "$PIPE_DATA/$T"
if [ ! -s dl/accns.txt ]; then
  CIK=$(python3 -c "import runpy;print(runpy.run_path('$CFG')['CIK'])")
  bash "$P/sec/dl.sh" "$CIK" dl
fi
python3 "$P/sec/build_gen.py" "$CFG" "$MASK" ${T}_a.xlsx | tail -1
bash "$P/sec/post.sh" $T | tail -1
python3 "$P/comum/shift_rows.py" ${T}_b_recalc.xlsx ${T}_c.xlsx Model 18 $N > /dev/null
python3 "$P/sec/gen_drivers.py" ${T}_c.xlsx ${T}_d.xlsx ${T}_a.xlsx.layout.json 18 $N "$SPEC" | tail -1
timeout 900 python3 "$P/comum/recalc.py" ${T}_d.xlsx ${T}_d_recalc.xlsx 2>&1 | grep -v javaldx || true
python3 "$P/comum/inject.py" ${T}_d.xlsx ${T}_d_recalc.xlsx ${T}_final.xlsx | tail -1
python3 "$P/comum/fix_panes.py" ${T}_final.xlsx
python3 "$P/sec/show.py" ${T}_final.xlsx 1 | grep -v -i warn
cp ${T}_final.xlsx "$REPO/Modelo_${NOME}_2T26.xlsx"
echo "gravado: Modelo_${NOME}_2T26.xlsx"
