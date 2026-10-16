#!/bin/bash
# Monta Modelo_Shell_2T26.xlsx a partir do Quarterly Databook do RI.
set -e
P="$(cd "$(dirname "$0")/../.." && pwd)"; REPO="$(cd "$P/.." && pwd)"
export PIPE_SEC="$P/sec"; export PIPE_DATA="$(mkdir -p "${PIPE_DATA:-$REPO/dados}" && cd "${PIPE_DATA:-$REPO/dados}" && pwd)"
MASK="$REPO/Modelo_2.0_proprio_Padrao_AJUSTADO-2028_corrigido.xlsx"
N=${1:-40}; NOME=${2:-Shell}
mkdir -p "$PIPE_DATA/shel/dl"
# Co() do gen_drivers exige um 20-F em accns.txt (só para o trimestre fiscal); fatos vazios.
if [ ! -s "$PIPE_DATA/shel/dl/accns.txt" ]; then
  echo '0001628280-26-017024 20-F 2026-03-12 dummy.xml 2025-12-31' > "$PIPE_DATA/shel/dl/accns.txt"
  python3 -I -c "import pickle; pickle.dump([], open('$PIPE_DATA/shel/dl/facts.pkl','wb'))"
fi
cd "$PIPE_DATA/shel"
python3 "$P/empresas/shel/build.py" "$MASK" shel_a.xlsx | tail -1
bash "$P/sec/post.sh" shel | tail -1
python3 "$P/comum/shift_rows.py" shel_b_recalc.xlsx shel_c.xlsx Model 18 $N > /dev/null
python3 "$P/sec/gen_drivers.py" shel_c.xlsx shel_d.xlsx shel_a.xlsx.layout.json 18 $N "$P/empresas/shel/spec.py" | tail -1
timeout 900 python3 "$P/comum/recalc.py" shel_d.xlsx shel_d_recalc.xlsx 2>&1 | grep -v javaldx || true
python3 "$P/comum/inject.py" shel_d.xlsx shel_d_recalc.xlsx shel_final.xlsx | tail -1
python3 "$P/comum/fix_panes.py" shel_final.xlsx
python3 "$P/sec/show.py" shel_final.xlsx 1 | grep -v -i warn
cp shel_final.xlsx "$REPO/Modelo_${NOME}_2T26.xlsx"
echo "gravado: Modelo_${NOME}_2T26.xlsx"
