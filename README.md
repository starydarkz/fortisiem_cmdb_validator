# FortiSIEM CMDB Validator

Herramienta de línea de comandos que se conecta a la API de **FortiSIEM** para revisar los equipos registrados en la **CMDB** (el inventario de dispositivos del SIEM) y decirte, de forma clara y en un archivo Excel, **cuáles están realmente enviando logs y cuáles no**.

Está pensada para equipos de SOC, CTI y administradores de FortiSIEM que necesitan auditar periódicamente el estado real de las integraciones, sin tener que revisar dispositivo por dispositivo desde la consola.

---

## ¿Qué problema resuelve?

En FortiSIEM es común que un dispositivo esté "aprobado" en la CMDB pero, por distintos motivos (cambio de IP, agente caído, firewall bloqueando el envío, etc.), **haya dejado de mandar eventos**. Detectar esto manualmente es lento. Esta herramienta:

1. Consulta la CMDB completa (o una lista de IPs que tú le indiques).
2. Le pregunta a FortiSIEM si cada equipo generó eventos en el rango de tiempo que definas.
3. Clasifica el tipo de evento y el protocolo de integración (Syslog, Agente, API, etc.).
4. Genera un reporte en Excel (`.xlsx`) con tres hojas: un resumen ejecutivo, el detalle por equipo y estadísticas.

## Características principales

- **Cobertura de logs**: porcentaje de equipos que sí están enviando eventos vs. los que no.
- **Detalle por dispositivo**: IP, hostname, estado de aprobación en la CMDB, tipo de dispositivo, validacion de envio de logs y detalles sobre que tipo de logs se reciben.
- **Clasificación automática de eventos**: identifica más de 100 tipos de integración (Windows, Linux, Fortinet, Cisco, Palo Alto, AWS, Azure, EDRs, bases de datos, etc.) y el protocolo usado.
- **Dos modos de validación**:
  - **Toda la CMDB** (`-xall`): analiza todos los equipos registrados en FortiSIEM.
  - **Lista específica** (`-i` / `--input`): valida solo las IPs que tú le pases (útil para auditar un subconjunto, un cliente, o un listado que te pidieron revisar).
- **Reporte visual en Excel**, listo para compartir con SOC Managers o clientes, sin necesidad de dar formato manualmente.

## ¿Cómo funciona? (en términos simples)

```
Tú ejecutas el programa  →  Se conecta a FortiSIEM (API REST)  →  Descarga el listado de la CMDB
                          →  Por cada equipo, pregunta "¿mandó eventos en las últimas X horas?"
                          →  Junta toda la información  →  Genera un archivo Excel con el resultado
```

No requiere instalar nada en FortiSIEM ni cambiar configuración del SIEM: usa la API REST que FortiSIEM ya expone.

---

## Requisitos

- Python 3.10 o superior (usa sintaxis moderna de type hints).
- Acceso de red hacia el **Supervisor de FortiSIEM** (puerto 443).
- Una cuenta de FortiSIEM con permisos de solo lectura sobre la CMDB y consultas de eventos (no requiere permisos de administración).
- Dependencias de Python (se instalan automáticamente, ver abajo):
  - `openpyxl` — generación del archivo Excel.
  - `httplib2` — consultas de eventos a FortiSIEM.
  - `tqdm` — barra de progreso en consola.

Nota: La ultima version testeada fue en la version de fortisiem v7.4.2, a partir de esta version en adelante la API funciona diferente y actualmente no es compatible.

---

## Instalación

Tienes dos formas de usar la herramienta: instalando Python directamente en tu equipo, o usando Docker (recomendado si no quieres instalar Python/dependencias localmente, o si vas a dejarlo corriendo en un servidor).

### Opción A — Instalación con Python (pip)

```bash
git clone https://github.com/starydarkz/Fortisiem_CMDB_Validator.git
cd Fortisiem_CMDB_Validator
pip3 install -r requirements.txt
```

Listo, ya puedes ejecutar el programa con `python3 fsmcmdbval.py -h` para ver opciones.

### Opción B — Instalación con Docker

Esta opción no requiere tener Python instalado en tu máquina. Solo necesitas Docker.

**1. Crea un `Dockerfile`** en la raíz del proyecto (junto a `fsmcmdbval.py`) con este contenido:

```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY fsmcmdbval.py .

# Carpeta donde se guardarán los reportes generados
RUN mkdir -p /app/output

ENTRYPOINT ["python3", "fsmcmdbval.py"]
```

**2. Construye la imagen:**

```bash
docker build -t fortisiem-cmdb-validator .
```

**3. Ejecútalo** (nota el volumen `-v`, que conecta una carpeta de tu equipo con la carpeta `/app/output` del contenedor, para que el Excel generado quede disponible fuera del contenedor):

```bash
mkdir -p reports

docker run --rm \
  -v "$(pwd)/reports:/app/output" \
  fortisiem-cmdb-validator \
  -u super/usuario -p "TuPassword" -s 192.168.1.10 \
  -o /app/output/reporte.xlsx -xall
```

Al terminar, encontrarás `reporte.xlsx` dentro de la carpeta `reports/` de tu equipo.

> 💡 Si prefieres no escribir el comando completo cada vez, más abajo (sección de Automatización) hay un ejemplo con `docker-compose` y con un programador de tareas para contenedores.

---

## Uso

```
python3 fsmcmdbval.py [opciones]
```

| Opción | Descripción |
|---|---|
| `-h`, `--help` | Muestra la ayuda y termina. |
| `-u`, `--user USER` | Usuario de FortiSIEM (ej: `super/usuario1`). |
| `-p`, `--passw PASSW` | Contraseña de la cuenta. |
| `-s`, `--siem SIEM` | IP o hostname del Supervisor de FortiSIEM. |
| `-o`, `--output OUTPUT` | *(Opcional)* Ruta del archivo de salida. Por defecto: `./fortisiem_report_devices.xlsx`. |
| `-t`, `--time TIME` | *(Opcional)* Rango de tiempo a analizar, en horas. Por defecto: `1`. |
| `-xall`, `--extractall` | Analiza **toda** la CMDB. |
| `-i`, `--input INPUT` | *(Opcional)* Valida solo IPs específicas: lista separada por comas, o ruta a un archivo `.csv`/`.txt`. |

> Debes usar **una** de las dos opciones de análisis: `-xall` (todo el inventario) o `-i` (una lista puntual de IPs). Si no se indica ninguna, el programa se detiene con un mensaje de error.

### Ejemplos de uso

**1. Analizar toda la CMDB, últimas 24 horas:**

```bash
python3 fsmcmdbval.py -u super/usuario -p "MiPassword" -s 192.168.1.10 -t 24 -o reporte_completo.xlsx -xall
```

**2. Validar solo un listado puntual de IPs (por ejemplo, equipos que te pidió el cliente revisar):**

```bash
python3 fsmcmdbval.py -u super/usuario -p "MiPassword" -s 192.168.1.10 -i "10.0.0.5,10.0.0.6,10.0.0.7"
```

**3. Validar IPs desde un archivo de texto (una IP por línea, o separadas por coma):**

```bash
python3 fsmcmdbval.py -u super/usuario -p "MiPassword" -s 192.168.1.10 -i ips.txt -o reporte_clientes.xlsx
```

> Nota: si una IP del listado (`-i`) no existe en la CMDB de FortiSIEM, igual aparecerá en el reporte marcada como **"No integrado en CMDB"**, para que sepas que ese equipo ni siquiera está registrado.

### Vista previa del reporte

![](./Example/2.png)
![](./Example/1.png)
![](./Example/3.png)

---

## Interpretando el reporte Excel

El archivo generado tiene 3 hojas:

1. **Resumen** — vista ejecutiva: rango de tiempo analizado, total de equipos, cuántos están enviando logs y cuántos no.
2. **Detalle** — tabla completa por equipo: IP, hostname, si está aprobado en la CMDB, tipo de dispositivo, estado de logs, tipos de eventos detectados y protocolo de integración.
3. **Estadísticas** — indicadores clave (KPI) y desglose de dispositivos por protocolo de integración.

---

## Notas de seguridad

- La contraseña se pasa como argumento de línea de comandos (`-p`). Esto significa que, mientras el proceso corre, puede ser visible para otros usuarios del mismo servidor (por ejemplo con `ps aux`). Se recomienda:
  - Usar una cuenta de FortiSIEM dedicada, de solo lectura, exclusiva para esta herramienta.
  - Evitar escribir la contraseña directamente en la terminal, en cambio optar por un archivo .env, crear variables de entorno temporales o hacer referencia desde algun otro archivo.
- La conexión hacia FortiSIEM se realiza sin validar el certificado SSL (`ssl._create_unverified_context`), algo común cuando el Supervisor usa un certificado autofirmado. Si tu entorno tiene un certificado válido, ten en cuenta que esta verificación está deshabilitada en el código actual.

---

## Bugs Conocidos

- Es posible que algunos equipos salgan como "Sin Logs" a pesar de que si se esten enviando eventos al FortiSIEM, este caso unicamente se ha visto cuando los unicos eventos que envia dicho equipo no estan siendo parseados correctamente y se les asigna el evento Unknow Event Type y no siempre sucede, parece un bug del propio fortisiem, por el momento esta en investigacion, asi que si un equipo sale como "Sin Logs", 100% no esta enviando logs o los eventos no se estan procesando correctamente.


## Licencia

GPL-3.0