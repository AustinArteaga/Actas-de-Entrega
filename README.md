# Acta de Entrega de Activos TI — Marriott S.A. Ecuador

Sistema web para registrar entregas de equipos informáticos, generar actas en Word (.docx) y mantener un inventario exportable.

## Archivos

```
marriott-ti/
├── index.html   # Estructura de la página
├── style.css    # Estilos visuales
├── app.js       # Toda la lógica (inventario, acta, Excel, JSON)
└── README.md
```

## Funcionalidades

- **Formulario** para registrar colaborador, equipo, accesorios y firmas
- **Acta en Word (.docx)** generada automáticamente al registrar
- **Inventario persistente** guardado en `localStorage` del navegador
- **Exportar Excel (.xlsx)** con todos los registros
- **Exportar / Importar JSON** para respaldar y restaurar el inventario
- **Eliminar registros** individuales desde el modal de inventario
- **Re-descargar actas** de registros anteriores

## Uso en GitHub Pages

1. Sube los 3 archivos (`index.html`, `style.css`, `app.js`) a un repositorio
2. Ve a **Settings → Pages → Branch: main → / (root)**
3. Accede a la URL que te genera GitHub Pages

## Librerías usadas (CDN, sin instalación)

- [docx.js 8.5](https://github.com/dolanmiu/docx) — Generación de archivos Word
- [SheetJS 0.18](https://sheetjs.com/) — Exportación a Excel
- [Google Fonts — Inter + DM Mono](https://fonts.google.com/)
