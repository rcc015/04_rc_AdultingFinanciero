# 04_rc_AdultingFinanciero

App de Google Apps Script para finanzas personales/compartidas.

## Archivos principales

- `code.gs`: backend de Apps Script.
- `index.html`: interfaz web.
- `appsscript.json`: manifiesto del proyecto.

## Conectar este repo con Apps Script usando `clasp`

1. Instala `clasp`:

```bash
npm install -g @google/clasp
```

2. Inicia sesión:

```bash
clasp login
```

3. Crea tu archivo local `.clasp.json` a partir de la plantilla:

```bash
cp .clasp.json.example .clasp.json
```

4. Edita `.clasp.json` y reemplaza `YOUR_SCRIPT_ID` por el `Script ID` de tu proyecto de Apps Script.

5. Sube el proyecto:

```bash
clasp push
```

6. Abre el proyecto en Apps Script:

```bash
clasp open
```

## Crear un proyecto nuevo de Apps Script

Si todavía no existe el script remoto:

```bash
clasp create --type standalone --title "AdultingFinanciero"
```

Después copia el `scriptId` generado a `.clasp.json` y ejecuta:

```bash
clasp push
```

## Desplegar como Web App

Desde Apps Script:

1. `Deploy` -> `New deployment`
2. Tipo `Web app`
3. Ejecutar como `Me`
4. Acceso según necesites

## Flujo de trabajo estable

Cada cambio importante debe cerrar con este flujo:

```bash
git status
git add .
git commit -m "mensaje claro del cambio"
git push
clasp push
```

Notas:

- `git push` respalda el código en GitHub.
- `clasp push` sincroniza el proyecto de Apps Script.
- Si el cambio debe verse en producción, además actualiza la implementación web desde Apps Script.

## Respaldo recomendado

### Código

- Haz commits pequeños y con mensajes claros.
- Empuja a GitHub después de cada cambio importante.
- Mantén `clasp push` como paso fijo después de cada `git push`.

### Datos

- Duplica o exporta periódicamente el Google Sheet que usa la app.
- Respaldo mínimo recomendado: antes de cambios grandes y al cierre de cada semana.
- Formato recomendado de exportación: Excel (`.xlsx`) o copia completa del Sheet en Drive.

### Metadata privada

- Guarda `Deployment URL`, `Script ID` y `Spreadsheet ID` en un archivo local no versionado.
- Usa la plantilla `PRIVATE_OPERATIONS.md.example` para crear tu copia local:

```bash
cp PRIVATE_OPERATIONS.md.example PRIVATE_OPERATIONS.md
```

## Scopes usados

Este proyecto usa servicios de:

- Google Sheets
- Gmail
- Mail
- Triggers
- HTML Service
- Session/User
