import logging
import sys

logger = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log', mode='w', encoding='utf-8'),  # 'w' для перезаписи файла
        logging.StreamHandler(sys.stdout)
    ]
)

sys.stderr = open('app.log', 'a')