"""
Face Processor - Detects, recorta y remueve el fondo de fotos de personas.
Usa OpenCV para detección facial y remove.bg API para remoción de fondo.
"""

import cv2
import numpy as np
import requests
import argparse
import os
from pathlib import Path

from dotenv import load_dotenv

# Cargar variables del archivo .env
load_dotenv()

# ─────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────

REMOVEBG_API_KEY = os.getenv("REMOVEBG_API_KEY")
PADDING_FACTOR   = 0.32                  # Padding alrededor de la cara (40%)
MIN_FACE_SIZE    = 30                   # Tamaño mínimo de cara en píxeles


def detect_faces(image: np.ndarray) -> list[tuple]:
    """Detecta caras en la imagen usando el clasificador Haar de OpenCV."""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    cascade_path = cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
    detector = cv2.CascadeClassifier(cascade_path)

    faces = detector.detectMultiScale(
        gray,
        scaleFactor=1.1,
        minNeighbors=5,
        minSize=(MIN_FACE_SIZE, MIN_FACE_SIZE),
    )
    return faces if len(faces) > 0 else []


def get_largest_face(faces: list) -> tuple:
    """Retorna la cara más grande detectada (asume que es la principal)."""
    return max(faces, key=lambda f: f[2] * f[3])


# ─────────────────────────────────────────
# RECORTE CON PADDING
# ─────────────────────────────────────────

def crop_face(image: np.ndarray, face: tuple, padding: float = PADDING_FACTOR) -> np.ndarray:
    h_img, w_img = image.shape[:2]
    x, y, w, h = face

    pad_x = int(w * padding)
    pad_top    = int(h * padding)        # padding arriba (cabello)
    pad_bottom = int(h * padding * 2.5)  # ← más espacio abajo (hombros)

    x1 = max(0, x - pad_x)
    y1 = max(0, y - pad_top)
    x2 = min(w_img, x + w + pad_x)
    y2 = min(h_img, y + h + pad_bottom)

    return image[y1:y2, x1:x2]

# ─────────────────────────────────────────
# REMOCIÓN DE FONDO (remove.bg)
# ─────────────────────────────────────────

def remove_background_removebg(image: np.ndarray, api_key: str) -> np.ndarray:
    """
    Envía la imagen a la API de remove.bg y retorna la imagen sin fondo (con canal alfa).
    """
    _, buffer = cv2.imencode(".png", image)
    img_bytes = buffer.tobytes()

    response = requests.post(
        "https://api.remove.bg/v1.0/removebg",
        files={"image_file": ("image.png", img_bytes, "image/png")},
        data={"size": "auto"},
        headers={"X-Api-Key": api_key},
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Error en remove.bg API [{response.status_code}]: {response.text}"
        )

    img_array = np.frombuffer(response.content, dtype=np.uint8)
    result = cv2.imdecode(img_array, cv2.IMREAD_UNCHANGED)

    if result is None:
        raise RuntimeError("No se pudo decodificar la respuesta de remove.bg.")

    return result


def remove_background_local(image: np.ndarray) -> np.ndarray:
    try:
        from rembg import remove
        from PIL import Image
        import io

        _, buffer = cv2.imencode(".png", image)
        pil_img = Image.open(io.BytesIO(buffer.tobytes()))
        result_pil = remove(pil_img)

        result_np = np.array(result_pil)
        result_bgra = cv2.cvtColor(result_np, cv2.COLOR_RGBA2BGRA)
        return result_bgra

    except ImportError as e:
        raise ImportError(f"Error importando rembg: {e}")  


# ─────────────────────────────────────────
# GUARDADO
# ─────────────────────────────────────────

def save_result(image: np.ndarray, output_path: str):
    """Guarda la imagen resultante. Soporta PNG (transparencia) y JPG."""
    ext = Path(output_path).suffix.lower()

    if ext == ".jpg" or ext == ".jpeg":
        # JPG no soporta transparencia → fondo blanco
        if image.shape[2] == 4:
            bgr = image[:, :, :3]
            alpha = image[:, :, 3:4] / 255.0
            white_bg = np.ones_like(bgr, dtype=np.uint8) * 255
            image = (bgr * alpha + white_bg * (1 - alpha)).astype(np.uint8)
        cv2.imwrite(output_path, image, [cv2.IMWRITE_JPEG_QUALITY, 95])
    else:
        ext = Path(output_path).suffix.lower()
        _, buffer = cv2.imencode(ext, image)
        with open(output_path, 'wb') as f:
            f.write(buffer.tobytes())

    print(f"✅ Resultado guardado en: {output_path}")


# ─────────────────────────────────────────
# PIPELINE PRINCIPAL
# ─────────────────────────────────────────

def process_image(
    input_path: str,
    output_path: str,
    mode: str = "local",
    api_key: str = REMOVEBG_API_KEY,
    padding: float = PADDING_FACTOR,
    skip_bg_removal: bool = False,
):
    print(f"\n📷 Procesando: {input_path}")

    # 1. Cargar imagen
    import numpy as np

    ruta_bytes = input_path.encode('utf-8') if isinstance(input_path, str) else input_path
    with open(input_path, 'rb') as f:
        buffer = np.frombuffer(f.read(), dtype=np.uint8)
    image = cv2.imdecode(buffer, cv2.IMREAD_COLOR)
    if image is None:
        raise FileNotFoundError(f"No se pudo cargar la imagen: {input_path}")
    
    print(f"   Tamaño original: {image.shape[1]}x{image.shape[0]}px")

    # 2. Detectar caras
    print("🔍 Detectando caras...")
    faces = detect_faces(image)
    if len(faces) == 0:
        raise ValueError("No se detectó ninguna cara en la imagen.")
    print(f"   {len(faces)} cara(s) encontrada(s). Usando la más grande.")

    # 3. Recortar alrededor de la cara principal
    face = get_largest_face(faces)
    cropped = crop_face(image, face, padding)
    print(f"   Recorte: {cropped.shape[1]}x{cropped.shape[0]}px (padding={padding:.0%})")

    if skip_bg_removal:
        # Guardar solo el recorte sin remover fondo
        save_result(cropped, output_path)
        return

    # 4. Remover fondo
    print(f"🎨 Removiendo fondo ({mode})...")
    if mode == "api":
        if api_key == "TU_API_KEY_AQUI":
            raise ValueError(
                "Configura tu API key de remove.bg en REMOVEBG_API_KEY "
                "o pásala con --api-key"
            )
        result = remove_background_removebg(cropped, api_key)
    elif mode == "local":
        result = remove_background_local(cropped)
    else:
        raise ValueError(f"Modo desconocido: {mode}. Usa 'api' o 'local'.")

    # 5. Guardar resultado
    save_result(result, output_path)


# ─────────────────────────────────────────
# CLI
# ─────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Recorta la cara de una foto y remueve el fondo."
    )
    parser.add_argument("input", help="Ruta a la imagen de entrada")
    parser.add_argument(
        "-o", "--output",
        help="Ruta de salida (default: <nombre>_processed.png)",
        default=None,
    )
    parser.add_argument(
        "--mode",
        choices=["api", "local"],
        default="api",
        help="'api' usa remove.bg | 'local' usa rembg (sin API key)",
    )
    parser.add_argument(
        "--api-key",
        default=None,
        help="API key de remove.bg (sobreescribe la constante del script)",
    )
    parser.add_argument(
        "--padding",
        type=float,
        default=PADDING_FACTOR,
        help=f"Padding alrededor de la cara (default: {PADDING_FACTOR})",
    )
    parser.add_argument(
        "--crop-only",
        action="store_true",
        help="Solo recortar la cara sin remover el fondo",
    )
    args = parser.parse_args()

    # Determinar ruta de salida
    if args.output is None:
        stem = Path(args.input).stem
        args.output = f"{stem}_processed.png"

    # API key
    api_key = args.api_key or REMOVEBG_API_KEY

    try:
        process_image(
            input_path=args.input,
            output_path=args.output,
            mode=args.mode,
            api_key=api_key,
            padding=args.padding,
            skip_bg_removal=args.crop_only,
        )
    except Exception as e:
        print(f"\n❌ Error: {e}")
        raise SystemExit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        input("\nPresiona Enter para cerrar...")
