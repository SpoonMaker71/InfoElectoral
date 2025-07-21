📘 Proyecto: InfoElectoral AccessVBA

![Revisado con Copilot](https://img.shields.io/badge/Revisado%20con-Copilot-00ADEF?logo=microsoft&logoColor=white&style=flat-square)
![Licencia](https://img.shields.io/badge/Licencia-GPL-blue)
![Tecnología](https://img.shields.io/badge/Microsoft%20Access-VBA-yellow)

🗳️ Proyecto Access VBA para importar y analizar resultados electorales por mesa en España.
Funciones reutilizables están nombradas en inglés; lógica específica del proyecto, en español.
Licencia: GNU GPL · código abierto y documentado para adaptaciones y mejoras.

🚀 Características principales
- 🔽 Importación automática de ficheros ZIP desde Infoelectoral.
- 📂 Descompresión nativa mediante Shell sin librerías externas.
- 🗃️ Lectura y decodificación de ficheros .DAT.
- 🧭 Navegación jerárquica: comunidad → provincia → municipio → mesa.
- 🧮 Cálculo de escaños con sistema D’Hondt.
- 📊 Exportación a Excel con formato, filtros y sumatorios.
- 🖥️ Formularios interactivos, barra de progreso, splash de inicio.

🔧 Requisitos
- Microsoft Access 2016 o superior
- Windows 10/11
- Permisos habilitados para macros
- Conexión a internet para descarga de ZIP
- Excel instalado (para exportación)

📂 Instalación
- Descarga o clona el repositorio:
git clone https://github.com/tuusuario/AccessElectoral.git
- Abre el archivo AccessDB.accdb.
- Inicio automático: El sistema lanza su formulario principal al abrir el archivo .accdb, siempre que las macros estén habilitadas. No requiere interacción directa con la ventana de Access.
- Sigue el flujo guiado para importar datos, navegar por niveles, consultar y exportar.

📤 Exportación a Excel
La función ToMSExcel permite generar un fichero .xlsx desde cualquier consulta SQL del sistema con:
- Cabecera destacada
- Totales automáticos
- Filtros en columnas
- Bordes, colores y formato profesional

📘 Créditos
Desarrollado por Juan Francisco Cucharero Cabezas
Inspirado por los microdatos públicos ofrecidos por Infoelectoral

📄 Licencia
Este proyecto se distribuye bajo la licencia GNU-GPL
Consulta el archivo LICENSE para más detalles.

🧠 Evaluación Técnica — Microsoft Copilot
Este repositorio ha sido acompañado y revisado por Microsoft Copilot, asistente de desarrollo inteligente.
Según su análisis técnico:
- El proyecto presenta una arquitectura modular, escalable y bien documentada, con separación clara entre lógica de negocio, funciones reutilizables y componentes visuales.
- Se destaca por una convención de nombres mixta (inglés para componentes reutilizables, español para lógica local), que aporta claridad y profesionalismo.
- La documentación embebida, el uso de clases, colecciones, y la guía editorial interna son señal de un proyecto con diseño sostenible y colaborativo.
- La estructura permite fácil mantenimiento, integración en otros sistemas y adaptación a nuevos procesos electorales.
En conjunto, InfoElectoral refleja un alto nivel profesional, propio de desarrollos con vocación institucional o comunitaria, y puede servir como base para sistemas más amplios de gestión electoral.

<svg viewBox="0 0 300 300" width="400" height="400" xmlns="http://www.w3.org/2000/svg">
  <style>
    .label { font: 12px sans-serif; fill: #333; }
    .value { font: bold 12px sans-serif; fill: #0078D7; }
    .line { stroke: #bbb; stroke-width: 1; }
    .shape { fill: rgba(0,120,215,0.2); stroke: #0078D7; stroke-width: 2; }
  </style>

  <!-- Center -->
  <circle cx="150" cy="150" r="2" fill="#0078D7"/>

  <!-- Axes -->
  <line x1="150" y1="150" x2="150" y2="30" class="line"/>
  <line x1="150" y1="150" x2="240" y2="65" class="line"/>
  <line x1="150" y1="150" x2="270" y2="150" class="line"/>
  <line x1="150" y1="150" x2="240" y2="235" class="line"/>
  <line x1="150" y1="150" x2="150" y2="270" class="line"/>
  <line x1="150" y1="150" x2="60" y2="235" class="line"/>
  <line x1="150" y1="150" x2="30" y2="150" class="line"/>
  <line x1="150" y1="150" x2="60" y2="65" class="line"/>

  <!-- Labels -->
  <text x="145" y="20" class="label" text-anchor="middle">Modularidad</text>
  <text x="250" y="65" class="label" text-anchor="middle">Documentación</text>
  <text x="280" y="150" class="label" text-anchor="middle">Automatización</text>
  <text x="250" y="250" class="label" text-anchor="middle">UX/UI</text>
  <text x="150" y="290" class="label" text-anchor="middle">Trazabilidad</text>
  <text x="50" y="250" class="label" text-anchor="middle">Exportación</text>
  <text x="20" y="150" class="label" text-anchor="middle">Análisis técnico</text>
  <text x="50" y="65" class="label" text-anchor="middle">Adaptabilidad</text>

  <!-- Valoración (radar shape) -->
  <polygon class="shape"
    points="
      150,50   <!-- Modularidad (10/10) -->
      225,85   <!-- Documentación (9/

Análisis realizado con el acompañamiento técnico de Microsoft Copilot — julio de 2025.

## ✨ Sobre el autor

**Juan Francisco Cucharero Cabezas** es desarrollador autodidacta con más de 27 años de experiencia. Comenzó su trayectoria en Visual Basic 6.0 y ha trabajado con lenguajes como C#, ASP 3.0, .NET, Java, Python, PHP, JavaScript y Appian. Aprendió por cuenta propia, complementando con cursos presenciales y formación online (incluidos recursos como YouTube).

Actualmente trabaja en **ATOS**, donde aplica su experiencia en entornos profesionales exigentes. Además, como hobby, diseña dispositivos físicos usando **Arduino** para enriquecer la experiencia en **Microsoft Flight Simulator**, combinando programación, electrónica y pasión por la simulación aérea.

Este proyecto refleja su compromiso con la transparencia
