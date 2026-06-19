# Porra Mundial 2026 🌍⚽

Web pública para seguir la porra del Mundial 2026 en tiempo real.

## Cómo publicar en GitHub Pages (gratis, 5 minutos)

### Primera vez

1. Crea una cuenta en [github.com](https://github.com) si no tienes
2. Crea un nuevo repositorio: botón **New** → nombre `porra-mundial-2026` → **Public** → **Create repository**
3. Sube los 3 archivos de esta carpeta (`index.html`, `data.json`, `update_data.py`):
   - En la página del repo, haz clic en **Add file → Upload files**
   - Arrastra los 3 archivos → **Commit changes**
4. Ve a **Settings → Pages → Source → Deploy from a branch → main → / (root) → Save**
5. En 1-2 minutos tu web estará en: `https://TU_USUARIO.github.io/porra-mundial-2026`

Comparte ese enlace con tus amigos — funciona en móvil y ordenador.

### Cada vez que actualices el Excel

**Opción A — Panel de Admin en la web (sin instalar nada):**
1. Abre la web
2. Haz clic en **↑ Admin** (arriba a la derecha)
3. PIN: `2026`
4. Sube el Excel → la web se actualiza al momento para ti
5. Para que todos lo vean, también tienes que actualizar `data.json` (opción B)

**Opción B — Actualizar data.json y subir a GitHub:**
1. Instala Python y openpyxl: `pip install openpyxl`
2. Ejecuta: `python update_data.py tu_excel.xlsx`
3. Se genera un `data.json` actualizado
4. Sube el nuevo `data.json` a GitHub:
   - En el repo, haz clic en `data.json` → botón **Edit (lápiz)** → pega el contenido → **Commit changes**
   - O usa GitHub Desktop para arrastrar y soltar

La web se actualiza automáticamente en 1-2 minutos.

## Archivos

- `index.html` — la aplicación web completa
- `data.json` — datos actuales (resultados + predicciones)
- `update_data.py` — script para regenerar data.json desde el Excel
