# Retail Omnicanal · Dashboard Ejecutivo y Auditoría de Metas (Reto A)

Proyecto final de análisis comercial, limpieza reproducible y visualización ejecutiva para la toma de decisiones sobre crecimiento, margen e inventario en retail omnicanal.

---

## 📊 Resumen Ejecutivo y Resultados Clave

- **Ventas Netas Limpias:** 780,933.87 BOB (2,422 transacciones válidas).
- **Meta Corporativa Total:** 55,970,111.13 BOB (545 combinaciones Mes · Región · Canal con meta presupuestada).
- **Cumplimiento Global Ponderado:** **1.39%** ($\frac{780,553.79}{55,970,111.13}$).
  > *Nota de Negocio:* El ratio de 1.39% refleja que la base entregada corresponde a una muestra transaccional frente a metas presupuestales macro de toda la corporación.
- **Margen Bruto Global:** 31.09% (Utilidad total: 242,764.55 BOB).
- **Tasa de Devolución:** 8.02% (Guardarraíl de control: ≤ 8.0%).
- **Auditoría Forense de Metas:** Se identificaron y corrigieron las 13 combinaciones con valores múltiples (por contaminación de tienda física o canal) y se documentó 1 combinación sin meta (Sep-2025 · Santa Cruz · Marketplace).

---

## 🚀 Guía de Despliegue Gratuito en la Nube

El proyecto está 100% configurado para desplegarse gratis en cualquiera de las siguientes plataformas:

### Opción 1: Streamlit Community Cloud (⭐ Recomendada · 100% Gratis y Permanente)

Streamlit Community Cloud es la plataforma oficial y óptima para aplicaciones interactivas de Streamlit:

1. **Subir el código a GitHub:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit - Retail Omnicanal Dashboard"
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/TU_REPOSITORIO.git
   git push -u origin main
   ```
2. **Desplegar en Streamlit:**
   - Entra a [share.streamlit.io](https://share.streamlit.io/) e inicia sesión con tu cuenta de GitHub.
   - Haz clic en **"New app"**.
   - Selecciona tu repositorio y rama (`main`).
   - En **Main file path**, escribe: `app.py` (o `dashboard_retail_omnicanal.py`).
   - Haz clic en **"Deploy!"**.
   - En menos de 2 minutos tu dashboard estará en vivo con URL pública (ej. `https://retail-omnicanal.streamlit.app`).

---

### Opción 2: Render (Free Web Service)

El repositorio incluye `Procfile` y `render.yaml` listos:

1. Crea una cuenta gratuita en [render.com](https://render.com/).
2. Haz clic en **New +** → **Web Service**.
3. Conecta tu repositorio de GitHub.
4. Configuración:
   - **Environment:** `Python 3`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `streamlit run app.py --server.port=$PORT --server.address=0.0.0.0`
5. Selecciona el plan **Free** y haz clic en **Create Web Service**.

---

### Opción 3: Hugging Face Spaces (Streamlit Gratuito)

1. Entra a [huggingface.co/spaces](https://huggingface.co/spaces) y crea un nuevo Space.
2. Selecciona **Space SDK:** `Streamlit`.
3. Sube los archivos del repositorio (o conecta tu GitHub).
4. Hugging Face detectará automáticamente `app.py` y `requirements.txt` y desplegará tu app en segundos.

---

### Opción 4: Vercel (Presentación Ejecutiva HTML)

El archivo `vercel.json` está configurado para publicar la presentación ejecutiva web:

1. Entra a [vercel.com](https://vercel.com/) e importa tu repositorio de GitHub.
2. Vercel detectará la configuración estática y publicará inmediatamente la `Presentacion_Ejecutiva_Retail_Omnicanal.html` en la raíz de tu dominio.

---

## 🛠️ Ejecución Local

Para correr el proyecto localmente en tu computadora:

```bash
# 1. Crear entorno virtual
python3 -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Generar entregables Excel y Presentación HTML
python3 generar_entregables.py

# 4. Iniciar el Dashboard interactivo
streamlit run app.py
```

---

## 📁 Estructura del Repositorio

- `app.py`: Punto de entrada principal para plataformas en la nube.
- `dashboard_retail_omnicanal.py`: Dashboard interactivo Streamlit con 6 vistas (Resumen Ejecutivo, Diagnóstico, Oportunidades, Insights, Plan de Acción y Calidad de Datos con Auditoría Forense).
- `preparar_datos.py`: Script de limpieza reproducible, cálculo de KPIs y auditoría de 14 combinaciones de metas.
- `generar_entregables.py`: Generador de `Caso_1_Retail_Omnicanal_Entrega.xlsx` y `Presentacion_Ejecutiva_Retail_Omnicanal.html`.
- `Caso_1_Retail_Omnicanal.xlsx`: Fuente de datos RAW original.
- `Caso_1_Retail_Omnicanal_Entrega.xlsx`: Libro de trabajo final procesado con todas las hojas documentadas (00 a 09 y 04b).
- `Presentacion_Ejecutiva_Retail_Omnicanal.html`: Presentación ejecutiva para Directorio.
- `requirements.txt`, `.streamlit/config.toml`, `Procfile`, `render.yaml`, `vercel.json`: Archivos de configuración para despliegue en la nube.
