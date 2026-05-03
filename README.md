# 📄 Extractor de Comprobantes con IA

Sistema de extracción automática de datos de **facturas, boletas y recibos peruanos** en formato PDF, con validación SUNAT y exportación a Excel profesional.

Convierte cientos de comprobantes en un reporte estructurado **en minutos**.

---

## ✨ Características

- 🤖 **Extracción inteligente con IA** (Google Gemini): identifica RUC, razón social, montos, ítems y fechas sin reglas hardcodeadas
- ✅ **Validación oficial SUNAT**: verifica RUC con dígito verificador y formato de series electrónicas
- 📊 **Excel profesional** con dos hojas: Resumen general y Detalle de ítems
- 📝 **Log de auditoría** con trazabilidad completa de cada procesamiento
- 📦 **Procesamiento por lotes**: cientos de PDFs en una sola ejecución
- 🎯 **Detección automática de tipo**: factura, boleta, recibo por honorarios, nota de crédito

## 🚀 Casos de uso

- Estudios contables que procesan comprobantes de múltiples clientes
- Empresas que digitalizan archivos físicos
- Departamentos de finanzas con alto volumen de facturas
- Auditorías que requieren consolidar comprobantes

## 📋 Requisitos

- Python 3.10+
- API Key gratuita de Google Gemini ([obtener aquí](https://aistudio.google.com/apikey))

## 🛠️ Instalación

```bash
git clone https://github.com/TU_USUARIO/extractor-comprobantes.git
cd extractor-comprobantes
python -m venv venv
.\venv\Scripts\Activate.ps1  # Windows
pip install -r requirements.txt
```

## ⚙️ Configuración

```bash
# Windows (PowerShell)
$env:GEMINI_API_KEY="tu_api_key_aqui"

# Linux/Mac
export GEMINI_API_KEY="tu_api_key_aqui"
```

## 💻 Uso

Coloca tus PDFs en la carpeta `pdfs/` y ejecuta:

```bash
python extractor.py
```

O especifica carpeta y salida personalizadas:

```bash
python extractor.py mis_facturas/ reporte_enero.xlsx
```

## 📊 Resultado

El sistema genera un Excel con dos hojas:

**Hoja "Resumen"** — un comprobante por fila:
| tipo_documento | ruc_emisor | razón_social | número | fecha | total | RUC válido |
|---|---|---|---|---|---|---|
| factura | 20123456789 | EMPRESA SAC | F001-00123 | 2026-01-15 | 1,180.00 | ✓ |

**Hoja "Detalle_Items"** — cada ítem en su propia fila para análisis detallado.

## 🔒 Privacidad y seguridad

- Los PDFs **nunca se almacenan** en servidores externos
- El procesamiento es local; solo el texto extraído se envía a la API de IA
- La API key se gestiona vía variables de entorno (no se hardcodea)

## 📄 Licencia

MIT — uso libre para proyectos personales y comerciales.

## 👤 Autor

Desarrollado por **Piero Barrantes** — disponible para proyectos de automatización contable y procesamiento de documentos.

📧 Contacto: [tu correo aquí]