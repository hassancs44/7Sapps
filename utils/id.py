from datetime import datetime
import random

def generate_custody_id():
    return f"CST-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(100,999)}"
