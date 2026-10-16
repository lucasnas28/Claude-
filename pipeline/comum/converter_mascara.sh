#!/bin/bash
# Passa modelos prontos da máscara antiga para a máscara vigente (to_new_mask + recálculo + painel + conferência).
# uso: pipeline/comum/converter_mascara.sh MASCARA SAIDA ARQUIVO:MOEDA ...   (ex.: Modelo_Oracle_2T26.xlsx:US$)
# variável: ULTIMO_TRI (padrão 2026-06-30)
W="$(cd "$(dirname "$0")" && pwd)"; MASK=$1; OUT=$2; shift 2; mkdir -p "$OUT"
for it in "$@"; do f=${it%%:*}; cur=${it##*:}; b=$(basename "$f" .xlsx)
  python3 "$W/to_new_mask.py" "$f" "$MASK" "$OUT/${b}_a.xlsx" "$cur" "${ULTIMO_TRI:-2026-06-30}" > "$OUT/${b}.log" 2>&1 || { echo "$b FALHA na conversão (ver $OUT/${b}.log)"; continue; }
  timeout 900 python3 "$W/recalc.py" "$OUT/${b}_a.xlsx" "$OUT/${b}_r.xlsx" > /dev/null 2>&1
  python3 "$W/inject.py" "$OUT/${b}_a.xlsx" "$OUT/${b}_r.xlsx" "$OUT/${b}.xlsx" > /dev/null 2>&1 && python3 "$W/fix_panes.py" "$OUT/${b}.xlsx" > /dev/null
  python3 "$W/check_modelo.py" "$OUT/${b}.xlsx"
  rm -f "$OUT/${b}_a.xlsx" "$OUT/${b}_r.xlsx"
done
