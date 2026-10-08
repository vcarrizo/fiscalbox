# FiscalBox

Sistema de gesti&oacute;n impositiva para estudios contables argentinos.

Aplicaci&oacute;n web standalone que se compila en un &uacute;nico archivo HTML autocontenido, sin necesidad de servidor backend. Todos los datos se almacenan en el navegador (localStorage).

## Funcionalidades

- **Liquidaci&oacute;n de IVA**: Importaci&oacute;n de comprobantes emitidos y recibidos desde archivos ARCA (XLS/XLSX/CSV). C&aacute;lculo autom&aacute;tico de d&eacute;bito y cr&eacute;dito fiscal con arrastre de saldo a favor.
- **Gesti&oacute;n multi-cliente**: Base de datos separada por cliente con soporte para Responsable Inscripto y Monotributista.
- **Monotributo**: Seguimiento de categor&iacute;a seg&uacute;n escalas vigentes.
- **IIBB Misiones**: Clasificaci&oacute;n autom&aacute;tica de facturaci&oacute;n por categor&iacute;a (Inscriptos, Consumidor Final, Exportaciones).
- **Retenciones y Percepciones**: Carga manual por per&iacute;odo.
- **Carga manual de comprobantes CAI**: Para comprobantes que no se importan desde ARCA.
- **Control de duplicados**: Detecci&oacute;n y gesti&oacute;n de comprobantes duplicados.
- **Conversi&oacute;n USD**: Detecci&oacute;n autom&aacute;tica de facturas en d&oacute;lares y conversi&oacute;n por tipo de cambio.
- **Contabilidad**: Plan de cuentas editable, asientos manuales, generaci&oacute;n autom&aacute;tica de asientos desde comprobantes, Balance de Sumas y Saldos.

## Requisitos

- Node.js 18+
- npm

## Instalaci&oacute;n

```bash
git clone https://github.com/tu-usuario/fiscalbox.git
cd fiscalbox
npm install
```

## Build

```bash
npm run build
```

Genera `dist/index.html` &mdash; un archivo HTML autocontenido (~1 MB) que se puede abrir directamente en el navegador o subir a cualquier hosting est&aacute;tico.

## Desarrollo

```bash
npm run dev
```

Inicia un servidor local en `http://localhost:3000`.

## Deploy en GitHub Pages

El repositorio incluye un workflow de GitHub Actions (`.github/workflows/deploy.yml`) que autom&aacute;ticamente compila y publica en GitHub Pages cada vez que se pushea a `main`.

Para activarlo:
1. En el repositorio de GitHub, ir a **Settings > Pages**
2. En **Source**, seleccionar **GitHub Actions**

## Stack

- React 18
- Recharts (gr&aacute;ficos)
- SheetJS/xlsx (lectura de archivos Excel)
- esbuild (bundler)
- Sin backend &mdash; 100% client-side

## Licencia

MIT
