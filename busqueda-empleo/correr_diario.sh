#!/bin/bash
# =============================================================================
#  correr_diario.sh — corre el buscador de empleos y avisa si hay NUEVAS
# =============================================================================
#  Lo dispara cron cada mañana (ver crontab). Pasos:
#    1) Ejecuta buscar_empleos.py (actualiza empleos_remotos.csv y resumen.json)
#    2) Lee cuántas vacantes NUEVAS salieron desde la última corrida
#    3) Si hay >=1 nueva, manda una notificación a Tide Commander
#  Todo queda registrado en cron.log (junto a este archivo).
# =============================================================================

# Carpeta donde vive este script (para que funcione sin importar desde dónde
# lo llame cron)
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR" || exit 1

# Registrar TODA la salida en registro.log. Si es interactivo (terminal) usamos
# tee para verlo también en pantalla; si lo dispara la Tarea de Windows (sin
# terminal), redirigimos directo al archivo — el 'tee' con process-substitution
# se perdía en ese contexto y dejaba el log vacío.
if [ -t 1 ]; then
  exec > >(tee -a "$DIR/registro.log") 2>&1
else
  exec >> "$DIR/registro.log" 2>&1
fi

AGENT_ID="x3294rs4"
PY=/usr/bin/python3
CURL=/usr/bin/curl

echo "===== $(date '+%Y-%m-%d %H:%M') ====="

# 0) Esperar a que haya red (al encender la laptop, la tarea puede arrancar
#    antes de que WSL tenga internet, lo que hacía fallar la corrida).
for i in $(seq 1 18); do
  "$CURL" -s --max-time 5 -o /dev/null https://remotive.com && break
  echo "esperando red... intento $i"
  sleep 5
done

# 1) Correr el buscador
"$PY" buscar_empleos.py

# 2) Leer el resumen (cuántas nuevas). Si algo falla, salir sin avisar.
[ -f resumen.json ] || { echo "sin resumen.json, no aviso"; exit 0; }

NUEVAS=$("$PY" -c "import json;print(json.load(open('resumen.json'))['nuevas'])" 2>/dev/null)
BILING=$("$PY" -c "import json;print(json.load(open('resumen.json'))['bilingues'])" 2>/dev/null)

# 3) Notificar solo si hay al menos una vacante nueva
if [ "${NUEVAS:-0}" -gt 0 ] 2>/dev/null; then
  MSG="${NUEVAS} vacantes nuevas (${BILING} piden espanol)"
  "$CURL" -s -X POST http://localhost:6200/api/notify \
    -H "Content-Type: application/json" \
    -d "{\"agentId\":\"${AGENT_ID}\",\"title\":\"Empleos remotos\",\"message\":\"${MSG}\"}" >/dev/null
  echo "avisado: ${MSG}"
else
  echo "sin vacantes nuevas hoy, no aviso"
fi
