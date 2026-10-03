# Verificación de equipos - COMTECH

La aplicación de escritorio genera correos HTML desde un Excel y los envía mediante Microsoft Graph. Cada usuario inicia sesión en la página oficial de Microsoft y los mensajes salen desde su propio buzón institucional. La aplicación no solicita ni almacena contraseñas.

## 1. Preparación del administrador de Microsoft 365

El administrador debe crear una App Registration en Microsoft Entra ID:

1. Ir a **Identity > Applications > App registrations > New registration**.
2. Usar un nombre como `Verificación de equipos COMTECH`.
3. Seleccionar **Accounts in this organizational directory only**.
4. En **Authentication**, agregar la plataforma **Mobile and desktop applications**.
5. Registrar esta URI de redirección:

   ```text
   http://localhost
   ```

6. Habilitar **Allow public client flows**.
7. En **API permissions**, agregar:

   ```text
   Microsoft Graph
   └── Delegated permissions
       └── Mail.Send
   ```

8. Otorgar consentimiento administrativo si la política de la organización lo requiere.
9. Entregar al operador:

   - `Directory (tenant) ID`;
   - `Application (client) ID`.

No se debe crear ni proporcionar un Client Secret. Esta es una aplicación de escritorio con inicio de sesión interactivo.

Si la aplicación empresarial tiene **Assignment required** activado, el administrador debe asignar los usuarios o grupos que podrán utilizarla.

## 2. Formato del Excel para escritorio

Para cargar equipos con los datos del CPU y del monitor, use estas seis columnas
en la primera fila de la primera hoja. Los encabezados repetidos deben ocupar
las posiciones indicadas:

| CPU | Marca y Modelo | Características del equipo | Monitores | Marca y Modelo | Características del equipo |
|---|---|---|---|---|---|
| MZ03WGLX | Lenovo ThinkCentre M70s Gen 6 | Ultra 7 265 / 32 GB / 1 TB / W11P | V90G0NDP | Lenovo ThinkVision T22i-30 | 22 pulgadas; DisplayPort / HDMI / VGA |

Cada fila genera un correo con ambas series, los modelos y las características
del CPU y el monitor. Las seis celdas son obligatorias; las filas incompletas
aparecen como inválidas en la vista previa y no se envían. Guarde las series
como texto para conservar ceros iniciales. También se acepta este formato en
CSV UTF-8.
El asunto indica la instalación del CPU y monitor en el Laboratorio 1 de
posgrado; el cuerpo incluye sus datos y la fecha y hora de instalación.

El formato anterior de inventario sigue disponible:

Para el inventario de CPU y monitores, la primera fila debe contener:

| Items | No. serie CPU | No. serie monitor | Responsable |
|---|---|---|---|
| 1 | MZ03WGK6 | V90G0NGP | Rollman |

Cada fila genera un correo individual con el ítem, ambas series, el responsable,
la fecha y la hora de preparación. El asunto usa la serie del CPU:
`Preparacion de CPU MZ03WGK6 para su respectivo uso.` Los cuatro campos son obligatorios. Guarde
las series como texto para conservar ceros iniciales. Se conservan las filas con
series repetidas: cada ítem se procesa por separado.

Los 40 registros están separados en tres archivos, con su numeración original:

- [equipos_rollman.xlsx](equipos_por_responsable/equipos_rollman.xlsx): 12 filas, ítems 1–12.
- [equipos_suyen.xlsx](equipos_por_responsable/equipos_suyen.xlsx): 14 filas, ítems 13–26.
- [equipos_dorian.xlsx](equipos_por_responsable/equipos_dorian.xlsx): 14 filas, ítems 27–40.

Para cada archivo, inicie sesión con la cuenta que debe enviar ese lote, seleccione
el Excel, pulse **Cargar y validar**, revise la vista previa y ejecute el lote.
Al terminar, pulse **Cerrar sesión**, inicie sesión con la siguiente cuenta y
cargue su Excel. El campo `Responsable` se incluye en el mensaje; el remitente
es la cuenta Microsoft conectada y el destino es el indicado en **Destinatario**.

La aplicación de escritorio admite directamente el formato de tickets:

| Codigo | Marca | Modelo | Serie | Lab |
|---|---|---|---|---|
| 07-1217-7001-83 | Dell | XPS | D7BLXM2 | Lab 2 |

Cada correo muestra el código, marca, modelo y serie del ticket. Guarde las
columnas `Codigo` y `Serie` como texto para conservar guiones o ceros iniciales.
La columna `Lab` es opcional; cuando está presente, genera el informe técnico
de diagnóstico con el asunto correspondiente al laboratorio. Se incluye
`diagnosticos_biblioteca_rusb.csv` con los 35 equipos proporcionados, listo para
seleccionar directamente desde el botón **Buscar** de la aplicación.

También se mantiene compatibilidad con el formato anterior. La primera hoja debe
contener estas columnas en la fila 1:

| NumeroSerie | Marcaymodelo | Caracteristicas |
|---|---|---|
| ABC0001 | Dell Latitude 5420 | Intel Core i5 / 16 GB RAM / SSD 512 GB |

Se recomienda guardar `NumeroSerie` como texto para conservar ceros iniciales. El programa registra las filas incompletas y omite números de serie repetidos dentro del mismo archivo.

## 3. Instalación

### CachyOS / Arch Linux

```bash
sudo pacman -S python python-tk
cd verificador_comtech
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

También puede ejecutar `./iniciar_linux.sh` después de instalar `python-tk`.

### Windows

Instale Python 3.11 o posterior desde python.org marcando `Add Python to PATH`. Después:

```powershell
cd verificador_comtech
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

También puede hacer doble clic en `iniciar_windows.bat`.

## 4. Uso

1. Introduzca el `Tenant ID` y `Client ID` proporcionados por el administrador.
2. Pulse **Iniciar sesión con Microsoft**.
3. El navegador abrirá la página oficial de Microsoft. Inicie sesión y complete MFA si se solicita.
4. La interfaz mostrará la cuenta conectada. Esa cuenta será el remitente.
5. Seleccione el Excel y configure la fecha y el rango horario.
6. Introduzca destinatario, CC y BCC.
7. Pulse **Cargar y validar** y revise la vista previa.
8. Mantenga activado **Modo prueba** y ejecute una simulación.
9. Desactive el modo prueba y confirme el envío real.

Para cambiar de remitente, pulse **Cerrar sesión** y después inicie sesión con otra cuenta autorizada de la organización.

El Tenant ID y Client ID se guardan en `datos/configuracion_oauth.json` para no tener que escribirlos nuevamente. No son secretos. Los tokens de Microsoft permanecen únicamente en memoria y desaparecen al cerrar el programa.

## 5. Envío mediante Microsoft Graph

La aplicación utiliza el permiso delegado `Mail.Send` y la operación:

```text
POST https://graph.microsoft.com/v1.0/me/sendMail
```

`/me` obliga a que el remitente sea el usuario que inició sesión. No es posible escribir otra dirección en el campo From para suplantar una cuenta. Los mensajes se guardan en Elementos enviados.

El destinatario puede ser interno o externo, sujeto a las políticas de transporte, seguridad y límites de Microsoft 365 de la organización.

## 6. Horarios

Las horas se distribuyen uniformemente dentro del rango indicado, excluyendo
el almuerzo de **12:00 a 13:00**, y se incluyen en el cuerpo del correo. A las
13:00 se pueden asignar horarios nuevamente. Si el rango está completamente
dentro del almuerzo, la aplicación solicita elegir otro rango.

Estos horarios no programan la hora real de salida. El envío real comienza al
pulsar el botón y respeta el delay configurado entre mensajes.

Con 600 registros entre 08:00 y 18:00 quedan nueve horas disponibles y se asigna
un horario de preparación cada 54 segundos, saltando el intervalo de almuerzo.

## 7. Historial y duplicados

La aplicación crea `datos/historial_envios.csv`. Un envío se considera duplicado cuando coinciden:

- cuenta remitente autenticada;
- destinatario;
- número de serie;
- fecha de verificación.

En el formato de CPU y monitores, la identidad del registro incluye el ítem,
la serie del CPU, la serie del monitor y el responsable, en lugar de usar solo
la serie. Así, las series repetidas en filas distintas no omiten correos. La
cuenta remitente también forma parte de la comprobación. El historial conserva
los registros anteriores y añade los campos del nuevo formato al registrar
el próximo lote.

Las simulaciones y los errores no bloquean un envío real posterior. No borre el historial si desea conservar la protección contra duplicados.

## 8. Errores comunes

- `AADSTS50011`: falta registrar correctamente `http://localhost` como redirect URI.
- `AADSTS65001` o consentimiento requerido: el administrador debe autorizar `Mail.Send`.
- `403 ErrorAccessDenied`: la cuenta o aplicación no tiene permiso para enviar.
- La ventana del navegador no regresa al programa: revise firewall local, redirect URI y que el puerto localhost temporal no esté bloqueado.
- El usuario no puede entrar: compruebe que pertenezca al tenant y esté asignado a la aplicación si **Assignment required** está activado.

## 9. Seguridad

- No introduzca contraseñas de Microsoft en la aplicación.
- Nunca comparta tokens OAuth ni archivos de sesiones.
- Mantenga el modo prueba activado hasta verificar destinatario, plantilla y cantidad de registros.
- El botón Cancelar detiene el lote después de la operación en curso.
# script_correos_uni_desktop
