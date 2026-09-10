from pathlib import Path
from urllib.request import urlopen

MODELS = {
    'face_detection_yunet_2023mar.onnx': 'https://github.com/opencv/opencv_zoo/raw/main/models/face_detection_yunet/face_detection_yunet_2023mar.onnx',
    'face_recognition_sface_2021dec.onnx': 'https://github.com/opencv/opencv_zoo/raw/main/models/face_recognition_sface/face_recognition_sface_2021dec.onnx',
}

out = Path(__file__).resolve().parent / 'models'
out.mkdir(exist_ok=True)
for name, url in MODELS.items():
    target = out / name
    if target.exists() and target.stat().st_size > 10000:
        print('OK', name)
        continue
    print('Downloading', name)
    with urlopen(url, timeout=120) as r, open(target, 'wb') as f:
        while True:
            chunk = r.read(1024 * 1024)
            if not chunk:
                break
            f.write(chunk)
    print('Saved', target, target.stat().st_size)
