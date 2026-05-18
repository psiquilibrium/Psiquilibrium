# Psiquilibrium Agenda

Aplicacion web para gestionar reservas de consultorios de Psiquilibrium.

## Arquitectura

- `index.html`: frontend estatico publicado en GitHub Pages.
- `Código.js`: backend de Google Apps Script.
- Google Sheets: persistencia de reservas, usuarios y bloqueos.

No cambiar `SCRIPT_URL` salvo que se migre explicitamente el despliegue de Apps Script.

## Desplegar frontend

Los cambios en `index.html` se publican con GitHub Pages al subirlos a `main`.

```bash
git add index.html
git commit -m "Describe el cambio"
git push origin main
```

## Desplegar backend

Los cambios en `Código.js` se suben al proyecto de Apps Script con `clasp`.

```bash
clasp status
clasp push
```

`clasp status` debe mostrar solo `Código.js` como archivo rastreado.

## Flujo recomendado

1. Editar archivos localmente.
2. Probar en produccion con una reserva o bloqueo de prueba.
3. Subir frontend con Git si cambió `index.html`.
4. Subir backend con `clasp push` si cambió `Código.js`.
5. Hacer commit de la configuracion/codigo relevante.

## Notas de seguridad

- No subir credenciales ni datos sensibles.
- No versionar `AGENTS.md`; es configuracion local de Codex.
- Mantener cambios pequenos y reversibles.
