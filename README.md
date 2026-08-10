# 🏆 Ventas SurOccidente — Dashboard Comercial

Dashboard interactivo para visualizar el cumplimiento de cuotas del equipo comercial de SurOccidente, con vista mensual y acumulada.

---

## 📁 Archivos del proyecto

```
index.html                    ← Portada (abrir este primero)
pages/mes.html                ← Vista detalle: mes actual
pages/acumulado.html          ← Vista detalle: acumulado del año
dashboard_ventas_style.css    ← Estilos y colores (editable)
dashboard_ventas.js           ← Lógica y lectura del Excel
ventas_occidente.xlsx         ← Datos del equipo (editable, tú lo actualizas)
```

> ⚠️ Todos los archivos deben mantener esta misma estructura de carpetas (el archivo `.xlsx` va en la raíz, junto a `index.html`, no dentro de `pages/`).

---

## 🚀 Cómo usar

### Opción A — VS Code (recomendado)
1. Abre la carpeta del proyecto en **Visual Studio Code**
2. Instala la extensión **Live Server**
3. Click derecho sobre `index.html` → **Open with Live Server**
4. Se abre en el navegador con los datos del Excel cargados automáticamente

### Opción B — Servidor local con Python
```bash
# En la carpeta del proyecto:
python -m http.server 8000
# Luego abre: http://localhost:8000/index.html
```

### Opción C — Doble click directo
Puede funcionar en algunos navegadores. Si los datos aparecen en cero o no cargan, usa la Opción A o B.

---

## 📊 Actualizar los datos

Edita el archivo `ventas_occidente.xlsx`. Tiene varias hojas (pestañas):

### Hoja 1 — Tabla principal (mes actual)

| Columna esperada | Descripción |
|---|---|
| `Pos` | Número de posición (opcional, se reordena solo) |
| `Comercial` / `Area` | Nombre del vendedor |
| `% de cumplimiento` | % de cuota cumplida en el mes |
| `Cuota` | Meta del mes |
| `Utilidad Bruta` | Utilidad generada en el mes |
| `Compromiso` | Compromiso pactado |
| `Falta` | Lo que falta para llegar a la meta |

> 💡 La tabla se **reordena automáticamente**: primero por % de cumplimiento, luego por utilidad bruta, luego por compromiso, y por último por qué tan cerca está de cumplir la meta ("Falta").

### Hoja 2 — Actividad Semanal
Actualmente **no se usa** en ninguna vista del dashboard. Puedes dejarla o quitarla, no afecta nada.

### Hoja "Dominio_Mes" y "Dominios_Acum"
Alimentan el ranking de dominios (mes actual y acumulado) que aparece en cada vista.

### Hoja "Cump_mes_a_mes"
Alimenta la matriz de "Histórico de Cumplimiento" (una fila por vendedor, una columna por mes).

---

## 🎨 Personalización de colores

En `dashboard_ventas_style.css`, sección `:root` (parte superior del archivo):

```css
--legend1-color : #d4b85a;
--legend2-color : #6a9ec0;
--legend3-color : #c47a7a;

--row-top-bg    : rgba(0, 60, 10, 0.75);
--row-warn-bg   : rgba(100, 70, 20, 0.70);
--row-elim-bg   : rgba(140, 10, 10, 0.80);
```

> Nota: estas variables de leyenda y de zonas de fila quedaron del diseño original y actualmente no se están usando en ninguna tabla visible. Se pueden limpiar más adelante si se desea.

---

## 🧭 Navegación entre páginas

- `index.html` → portada con 2 tarjetas: "Mes Actual" y "Acumulado"
- `pages/mes.html` → detalle del mes: tabla de vendedores, histórico de cumplimiento, ranking de dominios del mes, KPIs
- `pages/acumulado.html` → detalle acumulado: ranking anual, histórico de cumplimiento, ranking de dominios acumulado, KPIs

Cada página tiene un botón para volver a la portada o saltar a la otra vista.

---

## 🛠️ Tecnologías usadas

- HTML5 + CSS3 (sin frameworks)
- JavaScript vanilla
- [SheetJS (xlsx)](https://sheetjs.com/) — lectura del Excel desde el navegador
- Google Fonts — Bebas Neue + Rajdhani

---

## ❓ Problemas frecuentes

| Problema | Solución |
|---|---|
| Los datos aparecen en 0 o no cargan | Abre con Live Server o servidor local, no con doble click directo |
| El Excel no se carga | Verifica que `ventas_occidente.xlsx` esté en la raíz del proyecto, junto a `index.html` |
| Cambié el Excel y no se actualiza | Refresca el navegador con Ctrl+F5 (o Cmd+Shift+R en Mac) para forzar la recarga |
| Los colores de fila no cambian | Esa lógica de colores por estado está desactivada actualmente en el código |

---

## 🚢 Despliegue en Cloudflare Pages

- Build command: ninguno (no requiere build)
- Output directory: `/` (raíz del repositorio)
- Rama: `main`

Cada vez que subas cambios a GitHub en la rama `main`, Cloudflare Pages debería redesplegar automáticamente.

---

*Ventas SurOccidente · Powered by Mario Arango · 2026*
