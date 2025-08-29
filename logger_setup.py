import os
import logging
from datetime import datetime

from .config import SAVE_BASE_DIR

log_dir = os.path.join(SAVE_BASE_DIR, "vod_logs")
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f"log_{datetime.today().strftime('%Y%m%d')}.log")

logging.basicConfig(
    filename=log_file,
    filemode="w",  
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

logging.info("="*25 + f" [NEW RUN] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} " + "="*25)