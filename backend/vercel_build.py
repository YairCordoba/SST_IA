import os, sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
sys.path.append(str(BASE_DIR / "api"))

from server import app  # ← tu FastAPI original

def handler(event, context):
    from mangum import Mangum
    return Mangum(app)(event, context)
