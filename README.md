**Metabase Dashboard Extract**

Este proyecto automatiza la generación de informes turísticos personalizados a partir de dashboards de Metabase. El flujo completo está diseñado para integrarse con herramientas como n8n y Google Gemini, permitiendo crear reportes visuales y descriptivos en formato PDF basados en preguntas de usuarios.

**¿Cómo funciona?**

Formulario del cliente (Metabase):
Un cliente ingresa su email y una pregunta sobre los datos turísticos.

Automatización vía n8n:
Esa información llega a un webhook que activa este script.

Extracción de datos:
El script accede al dashboard de Metabase (con Selenium), captura los gráficos como imágenes y filtra los más relevantes con ayuda de un modelo de lenguaje de Google Gemini.

Generación del informe:
Se arma un documento Word con imágenes y descripciones automáticas de cada visualización, junto a una conclusión final generada con IA.

Exportación a PDF y envío:
El documento se convierte a PDF y se puede enviar al correo del cliente.

Tecnologías utilizadas
Python + FastAPI (API local para integrar con n8n)

Selenium (navegación y captura de dashboards)

Google Gemini API (descripciones y filtrado de visualizaciones)

LibreOffice + subprocess (conversión de Word a PDF)

n8n (automatización del flujo)

**Objetivo**

Permitir a municipios, oficinas de turismo o analistas generar informes claros y automatizados a partir de dashboards de datos, con una experiencia completamente integrada.
