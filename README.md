# 🏆 Copa Copilot — Dashboard de Posiciones

Dashboard interactivo con temática de fútbol / Mundial para visualizar el uso de **Microsoft Copilot** por área dentro de una organización.

---

## 📁 Archivos del proyecto

```
index.html   ← Página principal (abrir este)
copa_copilot_style.css        ← Estilos y colores (editable)
copa_copilot_app.js           ← Lógica y lectura del Excel
copa_copilot_datos.xlsx       ← Datos de las áreas (editable)
```

> ⚠️ Los 4 archivos deben estar **en la misma carpeta**.

---

## 🚀 Cómo usar

### Opción A — VS Code (recomendado)
1. Abre la carpeta en **Visual Studio Code**
2. Instala la extensión **Live Server**
3. Click derecho sobre `index.html` → **Open with Live Server**
4. Se abre automáticamente en el navegador con todos los datos

### Opción B — Servidor local Python
```bash
# En la carpeta del proyecto:
python -m http.server 8000
# Luego abre: http://localhost:8000/index.html
```

### Opción C — Doble click directo
Funciona en algunos navegadores (Edge, Firefox).  
Si los datos aparecen en cero, usa la Opción A o B.

---

## 📊 Actualizar los datos

Edita el archivo `copa_copilot_datos.xlsx`. Tiene **dos hojas**:

### Hoja 1 — `Copa Copilot` (tabla principal)

| Columna | Descripción | Ejemplo |
|---|---|---|
| `Pos` | Número de posición inicial | 1, 2, 3… |
| `Area` | Nombre del área o equipo | Ventas |
| `Usuarios_Activos_Pct` | % de usuarios activos (número entero) | 61 |
| `Interacciones` | Total de interacciones con Copilot | 890 |
| `Goles` | Puntuación / goles del equipo | 13 |
| `Estado` | Estado actual (ver tabla abajo) | Clasificado |

> 💡 La tabla se **reordena automáticamente** por Goles de mayor a menor.  
> En caso de empate, desempata por Interacciones.

### Hoja 2 — `Actividad Semanal` (gráfica de barras)

| Semana | Ventas | Marketing | Finanzas | RRHH | Admin |
|---|---|---|---|---|---|
| Semana 1 | 40 | 35 | 30 | 25 | 20 |
| Semana 2 | 48 | 40 | 38 | 30 | 22 |
| … | … | … | … | … | … |

> Las columnas deben coincidir con los nombres de las áreas de la Hoja 1.

---

## 🎨 Estados disponibles y colores

El color de fondo de cada fila se asigna automáticamente según la palabra clave en el campo **Estado**:

| Estado en Excel | Color de fila | Ícono | Descripción |
|---|---|---|---|
| `Clasificado` | 🟢 Verde | 🏆 | Área con buen desempeño |
| `Pre-clasificado` | 🟢 Verde | 🏆 | También detectado como clasificado |
| `En Carrera` | 🟡 Ámbar | 🔥 | En competencia, puede mejorar |
| `En Riesgo` | 🟡 Ámbar | 🔥 | También entra como ámbar |
| `Eliminado` | 🔴 Rojo | ❌ | Área con bajo desempeño |
| `Descalificado` | 🔴 Rojo | ❌ | También detectado como eliminado |

> La detección es flexible: no importan mayúsculas ni tildes.  
> Cualquier estado no reconocido aparece en **ámbar** por defecto.

---

## ✏️ Personalización

### Cambiar el nombre del cliente
En `index.html`, línea del título:
```html
<h1 class="title">Copa <span class="accent">Copilot</span> <span class="cliente">TuEmpresa</span></h1>
```

### Cambiar la temporada
```html
<p class="subtitle">⚽ &nbsp; TABLA DE POSICIONES · TEMPORADA 2026 &nbsp; ⚽</p>
```

### Cambiar colores de la leyenda
En `copa_copilot_style.css`, sección `:root` (líneas 18-20):
```css
--legend1-color : #d4b85a;   /* punto "50%+ Usuarios Activos"      */
--legend2-color : #6a9ec0;   /* punto "Interacciones con Copilot"  */
--legend3-color : #c47a7a;   /* punto "Zona de Eliminación"        */
```

### Cambiar colores de zonas de la tabla
```css
--row-top-bg    : rgba(0, 60, 10, 0.75);    /* verde — Clasificados  */
--row-warn-bg   : rgba(100, 70, 20, 0.70);  /* ámbar — En Carrera    */
--row-elim-bg   : rgba(140, 10, 10, 0.80);  /* rojo  — Eliminados    */
```

### Cambiar banderas del header
En `copa_copilot_dashboard.html`:
```html
<!-- Lado izquierdo -->
<span>🇨🇴</span><span>🇧🇷</span><span>🇦🇷</span>

<!-- Lado derecho -->
<span>🇫🇷</span><span>🇩🇪</span><span>🇪🇸</span>
```

---

## 📐 Lo que muestra el dashboard

| Sección | Descripción |
|---|---|
| **Tabla de posiciones** | Ranking ordenado por goles, con barras de % activos y badge de estado |
| **Actividad Semanal** | Gráfica de barras agrupadas por semana + línea de tendencia del líder |
| **Goles por Área** | Donut interactivo con balón giratorio |
| **Marcador General** | 4 KPIs: total goles, interacciones, líder y número de equipos |

---

## 🛠️ Tecnologías usadas

- HTML5 + CSS3 (sin frameworks)
- JavaScript vanilla
- [SheetJS (xlsx)](https://sheetjs.com/) — lectura del Excel desde el navegador
- [Google Fonts](https://fonts.google.com/) — Bebas Neue + Rajdhani
- SVG nativo para donut y línea de tendencia

---

## ❓ Problemas frecuentes

| Problema | Solución |
|---|---|
| Los datos aparecen en 0 | Abre con Live Server o servidor local (no doble click) |
| La gráfica semanal no aparece | Verifica que la Hoja 2 del Excel exista y tenga datos |
| Los colores de fila no cambian | Revisa la ortografía del Estado: debe contener "clasif" o "elim" |
| El Excel no se carga | Verifica que `copa_copilot_datos.xlsx` esté en la misma carpeta que el HTML |

---

*Copa Copilot · Powered by Microsoft Copilot · 2026*
