# Consejo Distrital de Archivos Medellín — Python v1

Aplicación web inicial para el Consejo Distrital de Archivos de Medellín.

Incluye el cuestionario de 21 preguntas, secciones A-E, matriz Q6, persistencia SQLite, cálculo IMA y panel administrativo básico.

## Instalación
```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```
Abrir http://127.0.0.1:5000

La base `consejo_cda.db` se crea automáticamente.

### Metodología
IMA = (puntaje obtenido / puntaje máximo aplicable) x 100. N/A se excluye del denominador. Preguntas evaluables: 5, 6, 7, 8, 9, 10, 14, 15, 16, 18 y 19.

Los rangos usados para los niveles corresponden al catálogo de la versión Excel de trabajo: 0–20, >20–40, >40–60, >60–80 y >80–100.
