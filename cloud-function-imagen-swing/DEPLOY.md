# Deploy de la Cloud Function — imagen comparativa del swing

No hace falta instalar nada en tu computadora (ni gcloud, ni Python). Todo se hace
desde el navegador, en la consola web de Google Cloud.

## 0. Aclaración: AI Studio pago ≠ esto

El plan pago de Google AI Studio te da más cuota para llamar a Gemini (lo que ya usás
en `codigo.gs`). Cloud Functions es un servicio distinto de Google Cloud — hay que
habilitarlo aparte, aunque uses el mismo proyecto.

## 1. Elegí o creá un proyecto de Google Cloud

1. Andá a https://console.cloud.google.com
2. Arriba a la izquierda, al lado de "Google Cloud", hay un selector de proyecto.
   Fijate si ya existe uno asociado a tu cuenta (puede haberse creado solo al activar
   AI Studio pago). Si preferís mantener todo separado y fácil de rastrear en costos,
   creá uno nuevo: "Proyecto nuevo" → nombre, por ejemplo `golfito-imagenes`.
3. Anotá el **ID del proyecto** (lo vas a ver en el selector).

## 2. Habilitá facturación

1. Menú ☰ (arriba a la izquierda) → **Facturación**.
2. Si el proyecto no tiene una cuenta de facturación vinculada, vinculá una
   (tarjeta de crédito/débito).
3. No te preocupes por el costo: Cloud Functions tiene una capa gratuita mensual muy
   generosa (2 millones de invocaciones). Con el volumen de Golfito esto no debería
   generar cargos reales, pero Google exige tener facturación habilitada para poder
   crear la función igual.

## 3. Creá la función

1. En el buscador de arriba, escribí **Cloud Functions** y entrá.
2. Click **"Crear función"**.
3. Configuración básica:
   - **Entorno**: 2nd gen
   - **Nombre de la función**: `generar-imagen-swing`
   - **Región**: `us-central1` (o la que uses habitualmente)
   - **Trigger**: HTTPS
   - **Autenticación**: "Permitir invocaciones no autenticadas" — la seguridad la
     maneja el secreto compartido (`FUNCTION_SECRET`) que configuramos en el paso 4,
     no el IAM de Google. Esto es igual al patrón que ya usás para los webhooks de
     Meta/MercadoPago en `codigo.gs` (`WEBHOOK_SECRET`).
4. Runtime, build, connections and security settings (desplegable, más abajo):
   - **Memoria**: 512 MB
   - **Tiempo de espera (timeout)**: 60s
   - **Variables de entorno de runtime** → **Agregar variable**:
     - Nombre: `FUNCTION_SECRET`
     - Valor: inventá un secreto largo y random (ej. 32+ caracteres). Guardalo,
       lo vas a necesitar en el paso 5. **No pongas nada trivial** — cualquiera que
       lo adivine podría invocar tu función y gastar tu cuota.
5. Click **Siguiente** (o "Next").

## 4. Pegá el código

1. **Entorno de ejecución (runtime)**: Python 3.12
2. **Punto de entrada (entry point)**: `generar_imagen_comparativa`
3. Editor insertado (inline editor):
   - Pestaña `main.py` → borrá el contenido de ejemplo y pegá el contenido de
     `cloud-function-imagen-swing/main.py` de este repo.
   - Pestaña `requirements.txt` → pegá el contenido de
     `cloud-function-imagen-swing/requirements.txt` de este repo.
4. Click **"Implementar" / "Deploy"**. Tarda 2-3 minutos.

## 5. Copiá la URL y configurala en Apps Script

1. Cuando termina el deploy, arriba de la página de la función vas a ver la
   **URL del trigger** (algo como
   `https://us-central1-tu-proyecto.cloudfunctions.net/generar-imagen-swing`).
   Copiala.
2. Andá al editor de Apps Script del proyecto Golfito → ⚙️ **Configuración del
   proyecto** → **Propiedades del script** → **Agregar propiedad del script**.
3. Agregá dos propiedades:
   - `IMG_CLOUD_FUNCTION_URL` = la URL que copiaste
   - `IMG_CLOUD_FUNCTION_SECRET` = el mismo secreto que pusiste como
     `FUNCTION_SECRET` en el paso 3 (tienen que ser idénticos, es lo que se
     comparan entre sí)

## 6. Probalo

Sin tocar WhatsApp: en el editor de Apps Script, seleccioná la función
`_testImagenComparativaSwing` en el desplegable de funciones (arriba), ejecutala
con dos parámetros desde el editor (`driveFileId` de un video de prueba ya subido a
tu carpeta `Golfito_Videos`, y tu propio número de WhatsApp) y revisá:
- El **Log de ejecución** en Apps Script (por si algo falló).
- Tu WhatsApp — deberías recibir la imagen "antes/después" como mensaje aparte,
  después del análisis de texto.

Si algo sale mal, el log de Apps Script te va a mostrar el error concreto
(Gemini, la Cloud Function, o el envío a WhatsApp) porque cada paso loguea el
suyo por separado.

## 7. En producción

Una vez que probaste y anduvo bien, el flujo real ya está enganchado: cualquier
alumno que mande un video de swing con ángulo "perfil" o "frontal" (no
"desconocido") va a recibir automáticamente la imagen comparativa como mensaje
extra, justo después del análisis de texto y antes del menú principal. Si ese
paso falla por cualquier motivo, el alumno igual recibe su análisis y el menú
con normalidad — solo se pierde el mensaje con la imagen.
