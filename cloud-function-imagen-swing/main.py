"""
Cloud Function HTTP: utilidades de imagen para el swing de golf. Dos acciones,
elegidas por el campo "accion" del body (JSON, POST):

accion "marcar_video" (default, comportamiento original):
{
  "video_url":  "https://www.googleapis.com/drive/v3/files/<id>?alt=media",
  "drive_token": "<bearer token OAuth con permiso de lectura sobre ese archivo de Drive>",
  "gemini_json": {
     "segundo_critico": 1.4,
     "error_bbox": [ymin, xmin, ymax, xmax],   // normalizado 0-1000
     "tipo_marca_error": "cruz" | "circulo",
     "texto_error": "...",
     "correccion_flecha_origen": [x, y],       // normalizado 0-1000
     "correccion_flecha_destino": [x, y],      // normalizado 0-1000
     "texto_correccion": "..."
  }
}

accion "pegar_logo" (pega el isotipo real de Golfito arriba a la derecha de una
imagen ya generada — no se le pide el logo a la IA, para que salga siempre
idéntico y nítido):
{
  "accion": "pegar_logo",
  "image_base64": "<base64 de la imagen sobre la que pegar el logo>",
  "logo_base64": "<base64 (o data URL) del isotipo, con transparencia>"
}

Header requerido en ambos casos: X-Auth-Secret, debe matchear la variable de
entorno FUNCTION_SECRET.

Salida (JSON) en ambos casos:
  200: {"ok": true, "image_base64": "...", "mime_type": "image/png"}
  4xx/5xx: {"ok": false, "error": "..."}
"""
import os
import io
import tempfile
import base64

import cv2
import numpy as np
import requests
import functions_framework
from PIL import Image, ImageDraw, ImageFont

MAX_VIDEO_BYTES = 60 * 1024 * 1024  # el bot ya rechaza videos de mas de 10s, esto es solo un techo de seguridad
DOWNLOAD_TIMEOUT = 20

ROJO = (0, 0, 230)     # BGR
VERDE = (60, 200, 60)
NEGRO = (0, 0, 0)
BLANCO = (255, 255, 255)


class ErrorEntrada(Exception):
    """Input invalido del caller (video_url/gemini_json faltante, etc). Mapea a HTTP 400."""


def _validar_secreto(request):
    secreto_esperado = os.environ.get("FUNCTION_SECRET")
    if not secreto_esperado:
        raise RuntimeError("FUNCTION_SECRET no configurado en la Cloud Function")
    secreto_recibido = request.headers.get("X-Auth-Secret", "")
    if secreto_recibido != secreto_esperado:
        raise PermissionError("Secreto invalido")


def _descargar_video(video_url, drive_token):
    headers = {"Authorization": f"Bearer {drive_token}"} if drive_token else {}
    try:
        resp = requests.get(video_url, headers=headers, stream=True, timeout=DOWNLOAD_TIMEOUT)
    except requests.RequestException as e:
        raise ErrorEntrada(f"No se pudo descargar el video: {e}")
    if resp.status_code != 200:
        raise ErrorEntrada(f"Descarga de video fallo con status {resp.status_code}")

    tmp = tempfile.NamedTemporaryFile(suffix=".mp4", delete=False)
    total = 0
    try:
        for chunk in resp.iter_content(chunk_size=1024 * 256):
            total += len(chunk)
            if total > MAX_VIDEO_BYTES:
                raise ErrorEntrada("El video supera el tamano maximo permitido")
            tmp.write(chunk)
    finally:
        tmp.close()
    if total == 0:
        os.unlink(tmp.name)
        raise ErrorEntrada("El video descargado esta vacio")
    return tmp.name


def _extraer_frame(video_path, segundo_critico):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        cap.release()
        raise ErrorEntrada("No se pudo abrir el video (formato no soportado o corrupto)")

    fps = cap.get(cv2.CAP_PROP_FPS) or 0
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)

    try:
        segundo = float(segundo_critico)
    except (TypeError, ValueError):
        segundo = None

    if fps > 0 and total_frames > 0 and segundo is not None:
        frame_idx = int(round(segundo * fps))
        frame_idx = max(0, min(total_frames - 1, frame_idx))
    elif total_frames > 0:
        # fallback: sin fps/segundo confiable, usamos el frame del medio
        frame_idx = total_frames // 2
    else:
        frame_idx = 0

    cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
    ok, frame = cap.read()
    if not ok or frame is None:
        # fallback: si el frame puntual no se pudo leer, probamos el primero
        cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
        ok, frame = cap.read()
    cap.release()
    if not ok or frame is None:
        raise ErrorEntrada("No se pudo extraer ningun frame del video")
    return frame


def _clamp(v, lo, hi):
    return max(lo, min(hi, v))


def _parse_bbox(bbox, w, h):
    """bbox normalizado 0-1000 [ymin,xmin,ymax,xmax] -> (x1,y1,x2,y2) en pixeles, o None si invalido."""
    try:
        ymin, xmin, ymax, xmax = [float(v) for v in bbox]
    except (TypeError, ValueError, IndexError):
        return None
    x1 = _clamp(xmin / 1000.0 * w, 0, w - 1)
    y1 = _clamp(ymin / 1000.0 * h, 0, h - 1)
    x2 = _clamp(xmax / 1000.0 * w, 0, w - 1)
    y2 = _clamp(ymax / 1000.0 * h, 0, h - 1)
    if x2 <= x1:
        x2 = min(w - 1, x1 + 1)
    if y2 <= y1:
        y2 = min(h - 1, y1 + 1)
    return int(x1), int(y1), int(x2), int(y2)


def _parse_punto(punto, w, h):
    try:
        x, y = [float(v) for v in punto]
    except (TypeError, ValueError, IndexError):
        return None
    return int(_clamp(x / 1000.0 * w, 0, w - 1)), int(_clamp(y / 1000.0 * h, 0, h - 1))


def _bbox_default(w, h):
    # caja centrada, 40% del frame -> siempre hay algo razonable para marcar aunque Gemini no haya devuelto nada usable
    x1, y1 = int(w * 0.3), int(h * 0.3)
    x2, y2 = int(w * 0.7), int(h * 0.7)
    return x1, y1, x2, y2


def _linea_con_contorno(img, p1, p2, color, grosor):
    cv2.line(img, p1, p2, NEGRO, grosor + 3, cv2.LINE_AA)
    cv2.line(img, p1, p2, color, grosor, cv2.LINE_AA)


def _dibujar_marca_error(frame, bbox, tipo):
    x1, y1, x2, y2 = bbox
    cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
    radio = max((x2 - x1), (y2 - y1)) // 2 + 8

    if tipo == "cruz":
        _linea_con_contorno(frame, (x1, y1), (x2, y2), ROJO, 4)
        _linea_con_contorno(frame, (x1, y2), (x2, y1), ROJO, 4)
    else:
        # "circulo" y cualquier tipo no reconocido (ej. futuro "lineas") caen aca por defecto
        cv2.circle(frame, (cx, cy), radio, NEGRO, 7, cv2.LINE_AA)
        cv2.circle(frame, (cx, cy), radio, ROJO, 4, cv2.LINE_AA)
    return frame


def _dibujar_flecha(frame, origen, destino):
    _linea_con_contorno(frame, origen, destino, VERDE, 5)
    cv2.arrowedLine(frame, origen, destino, VERDE, 5, cv2.LINE_AA, tipLength=0.3)
    return frame


# La fuente por defecto de Pillow (ImageFont.load_default) NO tiene tildes ni ñ — para texto en
# español hace falta una fuente con cobertura Unicode real. Como el editor inline de Cloud Run
# solo permite pegar código como texto (no subir archivos binarios), la descargamos una vez por
# instancia y la cacheamos en /tmp para los requests siguientes en la misma instancia.
FONT_URL = "https://raw.githubusercontent.com/matplotlib/matplotlib/main/lib/matplotlib/mpl-data/fonts/ttf/DejaVuSans.ttf"
FONT_PATH = "/tmp/DejaVuSans.ttf"


def _asegurar_fuente_descargada():
    if os.path.exists(FONT_PATH) and os.path.getsize(FONT_PATH) > 0:
        return
    resp = requests.get(FONT_URL, timeout=15)
    resp.raise_for_status()
    with open(FONT_PATH, "wb") as f:
        f.write(resp.content)


def _cargar_fuente(size):
    try:
        _asegurar_fuente_descargada()
        return ImageFont.truetype(FONT_PATH, size)
    except Exception:
        # Fallback si no hay red o la descarga falla: al menos no rompe el flujo,
        # aunque sin tildes/ñ.
        try:
            return ImageFont.load_default(size=size)
        except TypeError:
            return ImageFont.load_default()


def _envolver_texto(draw, texto, font, max_width):
    palabras = texto.split()
    lineas = []
    actual = ""
    for palabra in palabras:
        candidato = (actual + " " + palabra).strip()
        if draw.textlength(candidato, font=font) <= max_width or not actual:
            actual = candidato
        else:
            lineas.append(actual)
            actual = palabra
    if actual:
        lineas.append(actual)
    return lineas[:3] or [""]  # evita que una respuesta muy larga de Gemini rompa el layout


def _agregar_banda_texto(frame, texto, color_acento_bgr):
    if not texto:
        texto = ""
    h, w = frame.shape[:2]
    tam_fuente = max(16, int(w / 42))
    font = _cargar_fuente(tam_fuente)

    medidor = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    lineas = _envolver_texto(medidor, texto, font, w - 28)

    alto_linea = tam_fuente + 14
    alto_banda = alto_linea * len(lineas) + 16

    overlay = frame.copy()
    cv2.rectangle(overlay, (0, h - alto_banda), (w, h), NEGRO, -1)
    frame[:] = cv2.addWeighted(overlay, 0.65, frame, 0.35, 0)
    cv2.rectangle(frame, (0, h - alto_banda), (6, h), color_acento_bgr, -1)

    # Pillow usa RGB y OpenCV BGR — convertimos solo para dibujar el texto (con tildes/ñ, que
    # cv2.putText no soporta) y volvemos a BGR para seguir componiendo con OpenCV.
    pil_img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(pil_img)
    y = h - alto_banda + 8
    for linea in lineas:
        draw.text((14, y), linea, font=font, fill=(255, 255, 255))
        y += alto_linea
    frame[:] = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    return frame


def _combinar_lado_a_lado(img_izq, img_der, label_izq, label_der):
    h, w = img_izq.shape[:2]
    alto_header = 46
    divisor_ancho = 4

    canvas = np.zeros((h + alto_header, w * 2 + divisor_ancho, 3), dtype=np.uint8)
    canvas[:] = (25, 25, 25)

    font = cv2.FONT_HERSHEY_SIMPLEX
    for texto, x_off in ((label_izq, 0), (label_der, w + divisor_ancho)):
        (tw, th), _ = cv2.getTextSize(texto, font, 0.8, 2)
        cv2.putText(canvas, texto, (x_off + (w - tw) // 2, 32), font, 0.8, BLANCO, 2, cv2.LINE_AA)

    canvas[alto_header:alto_header + h, 0:w] = img_izq
    canvas[alto_header:alto_header + h, w + divisor_ancho:w * 2 + divisor_ancho] = img_der
    return canvas


def _generar_imagen(frame_original, gemini_json):
    h, w = frame_original.shape[:2]

    bbox = _parse_bbox(gemini_json.get("error_bbox"), w, h) or _bbox_default(w, h)
    tipo_marca = gemini_json.get("tipo_marca_error") if gemini_json.get("tipo_marca_error") in ("cruz", "circulo") else "circulo"
    texto_error = gemini_json.get("texto_error") or "Error detectado en el swing"

    origen = _parse_punto(gemini_json.get("correccion_flecha_origen"), w, h)
    destino = _parse_punto(gemini_json.get("correccion_flecha_destino"), w, h)
    if not origen or not destino:
        cx, cy = (bbox[0] + bbox[2]) // 2, (bbox[1] + bbox[3]) // 2
        origen, destino = (cx, cy), (cx, max(0, cy - int(h * 0.15)))
    texto_correccion = gemini_json.get("texto_correccion") or "Correccion sugerida"

    img_antes = frame_original.copy()
    _dibujar_marca_error(img_antes, bbox, tipo_marca)
    _agregar_banda_texto(img_antes, texto_error, ROJO)

    img_despues = frame_original.copy()
    _dibujar_flecha(img_despues, origen, destino)
    _agregar_banda_texto(img_despues, texto_correccion, VERDE)

    return _combinar_lado_a_lado(img_antes, img_despues, "ANTES", "DESPUES")


def _decodificar_imagen_base64(b64_str, nombre_campo):
    if not b64_str:
        raise ErrorEntrada(f"Falta {nombre_campo}")
    b64_limpio = b64_str.split(",", 1)[-1]  # por si viene con prefijo data:image/png;base64,
    try:
        raw = base64.b64decode(b64_limpio)
    except Exception as e:
        raise ErrorEntrada(f"{nombre_campo} invalido (no es base64 valido): {e}")
    try:
        return Image.open(io.BytesIO(raw)).convert("RGBA")
    except Exception as e:
        raise ErrorEntrada(f"{nombre_campo} invalido (no se pudo decodificar como imagen): {e}")


def _pegar_logo(imagen, logo_base64):
    if not logo_base64:
        return imagen
    logo = _decodificar_imagen_base64(logo_base64, "logo_base64")

    w, _h = imagen.size
    margen = int(w * 0.03)
    ancho_logo = int(w * 0.14)
    ratio = ancho_logo / logo.width
    alto_logo = max(1, int(logo.height * ratio))
    logo_resized = logo.resize((ancho_logo, alto_logo), Image.LANCZOS)

    resultado = imagen.copy()
    x = w - ancho_logo - margen
    y = margen
    resultado.paste(logo_resized, (x, y), logo_resized)
    return resultado


def _accion_pegar_logo(body):
    imagen = _decodificar_imagen_base64(body.get("image_base64"), "image_base64")
    imagen_con_logo = _pegar_logo(imagen, body.get("logo_base64"))
    buf = io.BytesIO()
    imagen_con_logo.save(buf, format="PNG")
    return {"ok": True, "image_base64": base64.b64encode(buf.getvalue()).decode("ascii"), "mime_type": "image/png"}, 200


def _accion_marcar_video(body):
    tmp_video_path = None
    try:
        video_url = body.get("video_url")
        gemini_json = body.get("gemini_json") or {}
        drive_token = body.get("drive_token")

        if not video_url:
            raise ErrorEntrada("Falta video_url")

        tmp_video_path = _descargar_video(video_url, drive_token)
        frame = _extraer_frame(tmp_video_path, gemini_json.get("segundo_critico"))
        imagen_final = _generar_imagen(frame, gemini_json)

        # PNG sin compresion con perdida (no usamos JPEG a proposito)
        ok, buf = cv2.imencode(".png", imagen_final, [cv2.IMWRITE_PNG_COMPRESSION, 3])
        if not ok:
            raise RuntimeError("No se pudo codificar la imagen final")

        return {
            "ok": True,
            "image_base64": base64.b64encode(buf.tobytes()).decode("ascii"),
            "mime_type": "image/png",
        }, 200
    finally:
        if tmp_video_path and os.path.exists(tmp_video_path):
            os.unlink(tmp_video_path)


@functions_framework.http
def generar_imagen_comparativa(request):
    try:
        _validar_secreto(request)
        body = request.get_json(silent=True) or {}
        accion = body.get("accion") or "marcar_video"

        if accion == "pegar_logo":
            return _accion_pegar_logo(body)
        elif accion == "marcar_video":
            return _accion_marcar_video(body)
        else:
            raise ErrorEntrada(f"accion desconocida: {accion}")

    except PermissionError as e:
        return {"ok": False, "error": str(e)}, 401
    except ErrorEntrada as e:
        return {"ok": False, "error": str(e)}, 400
    except Exception as e:
        return {"ok": False, "error": f"Error interno: {e}"}, 500
