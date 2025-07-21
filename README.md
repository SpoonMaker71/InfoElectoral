📘 Proyecto: InfoElectoral AccessVBA

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

## ✨ Sobre el autor

**Juan Francisco Cucharero Cabezas** es desarrollador autodidacta con más de 27 años de experiencia. Comenzó su trayectoria en Visual Basic 6.0 y ha trabajado con lenguajes como C#, ASP 3.0, .NET, Java, Python, PHP, JavaScript y Appian. Aprendió por cuenta propia, complementando con cursos presenciales y formación online (incluidos recursos como YouTube).

Actualmente trabaja en **ATOS**, donde aplica su experiencia en entornos profesionales exigentes. Además, como hobby, diseña dispositivos físicos usando **Arduino** para enriquecer la experiencia en **Microsoft Flight Simulator**, combinando programación, electrónica y pasión por la simulación aérea.

Este proyecto refleja su compromiso con la transparencia
