#!/usr/bin/env bash
# Inicia el dashboard con el entorno virtual correcto.
set -e
cd "$(dirname "$0")"

if [[ ! -d venv ]]; then
  echo "Error: no existe la carpeta venv. Crea el entorno e instala dependencias primero."
  exit 1
fi

source venv/bin/activate

# Generar Excel de entrega si falta
if [[ ! -f Caso_1_Retail_Omnicanal_Entrega.xlsx ]]; then
  echo "Generando Caso_1_Retail_Omnicanal_Entrega.xlsx..."
  python generar_entregables.py
fi

echo ""
echo "Dashboard disponible en: http://localhost:8501"
echo "Presiona Ctrl+C para detener."
echo ""

exec streamlit run dashboard_retail_omnicanal.py --server.headless true
