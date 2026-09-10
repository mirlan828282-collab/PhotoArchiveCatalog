import os
import sys
from pathlib import Path

try:
    import cv2
    import numpy as np
except Exception:
    cv2 = None
    np = None


def resource_dir():
    if getattr(sys, 'frozen', False):
        return Path(getattr(sys, '_MEIPASS', Path(sys.executable).parent))
    return Path(__file__).resolve().parent


MODEL_DIR = resource_dir() / 'models'
DETECTOR_MODEL = MODEL_DIR / 'face_detection_yunet_2023mar.onnx'
RECOGNIZER_MODEL = MODEL_DIR / 'face_recognition_sface_2021dec.onnx'


class FaceEngine:
    def __init__(self):
        self.available = False
        self.error = ''
        self.detector = None
        self.recognizer = None
        if cv2 is None or np is None:
            self.error = 'OpenCV не установлен.'
            return
        if not DETECTOR_MODEL.exists() or not RECOGNIZER_MODEL.exists():
            self.error = 'Модели распознавания лиц не найдены в папке models.'
            return
        try:
            self.detector = cv2.FaceDetectorYN_create(str(DETECTOR_MODEL), '', (320, 320), 0.75, 0.3, 5000)
            self.recognizer = cv2.FaceRecognizerSF_create(str(RECOGNIZER_MODEL), '')
            self.available = True
        except Exception as exc:
            self.error = str(exc)

    def embedding(self, image_path):
        if not self.available:
            raise RuntimeError(self.error or 'Модуль поиска по лицу недоступен.')
        img = cv2.imread(str(image_path))
        if img is None:
            raise ValueError('Не удалось открыть изображение.')
        h, w = img.shape[:2]
        self.detector.setInputSize((w, h))
        _, faces = self.detector.detect(img)
        if faces is None or len(faces) == 0:
            raise ValueError('Лицо на фотографии не найдено.')
        # Берём самое крупное лицо как основное.
        face = max(faces, key=lambda f: float(f[2] * f[3]))
        aligned = self.recognizer.alignCrop(img, face)
        feat = self.recognizer.feature(aligned).flatten().astype('float32')
        norm = float(np.linalg.norm(feat))
        if norm > 0:
            feat /= norm
        return feat

    @staticmethod
    def cosine(a, b):
        if np is None:
            return 0.0
        a = np.asarray(a, dtype='float32')
        b = np.asarray(b, dtype='float32')
        den = float(np.linalg.norm(a) * np.linalg.norm(b))
        return float(np.dot(a, b) / den) if den else 0.0
