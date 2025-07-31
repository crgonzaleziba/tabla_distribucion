# Prueba Julio

Este proyecto automatiza la generación y análisis de reportes de rentabilidad y distribución de carteras, combinando datos de archivos Excel y consultas SQL.

## Requisitos

- Python 3.8+
- Paquetes:
  - pandas
  - openpyxl
  - sqlalchemy
  - python-dotenv
  - pymssql (para conexión SQL Server)

## Instalación

1. Clona el repositorio o copia los archivos en tu carpeta de trabajo.
2. Instala los paquetes necesarios:
   ```sh
   pip install pandas openpyxl sqlalchemy python-dotenv pymssql
   ```
3. Crea un archivo `.env` en la raíz del proyecto con tus credenciales SQL:
   ```
   SQL_SERVER=tu_servidor
   SQL_PORT=1433
   SQL_USER=tu_usuario
   SQL_PASSWORD=tu_contraseña
   SQL_DATABASE=tu_base
   SQL_SCHEMA=dbo
   ```
4. Coloca los archivos Excel requeridos en la carpeta del proyecto:
   - Clasificación de funcionarios por Unidad de negocio.xlsx
   - 2025 Información Rentabilidad Carteras - Analítica.xlsx

## Uso

Ejecuta el script principal:
```sh
python pruebajulio.py
```

El script generará y actualizará el archivo `Resultado_distribucion.xlsx` con las hojas y cálculos solicitados.

## Funcionalidades principales
- Procesamiento y limpieza de datos de funcionarios y carteras.
- Generación de reportes y hojas Excel con fórmulas y porcentajes.
- Consulta y extracción de datos desde SQL Server.
- Cálculo de distribuciones y combinaciones para análisis de ciudad y comercio.
- Uso de variables de entorno para credenciales seguras.

## Seguridad
- Las credenciales SQL se almacenan en `.env` y nunca en el código fuente.
- No subas archivos `.env` ni Excel con datos personales a repositorios públicos.

## Contacto
Para soporte o dudas, contacta a tu equipo de analítica o al desarrollador responsable.
# tabla_distribucion
