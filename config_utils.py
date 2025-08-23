import os
import json
from cryptography.fernet import Fernet

CONFIG_PATH = os.path.expanduser("~/.myapp_config.enc")
KEY_PATH = os.path.expanduser("~/.myapp_key.key")

def generate_key():
    """Generate and save a key for encryption"""
    key = Fernet.generate_key()
    with open(KEY_PATH, "wb") as f:
        f.write(key)
    return key

def load_key():
    """Load the encryption key or generate if missing"""
    if not os.path.exists(KEY_PATH):
        return generate_key()
    with open(KEY_PATH, "rb") as f:
        return f.read()

def encrypt_config(data: dict, key: bytes) -> bytes:
    f = Fernet(key)
    json_data = json.dumps(data).encode()
    return f.encrypt(json_data)

def decrypt_config(enc_data: bytes, key: bytes) -> dict:
    f = Fernet(key)
    decrypted = f.decrypt(enc_data)
    return json.loads(decrypted.decode())

def save_config(data: dict):
    key = load_key()
    enc = encrypt_config(data, key)
    with open(CONFIG_PATH, "wb") as f:
        f.write(enc)

def load_config() -> dict:
    if not os.path.exists(CONFIG_PATH):
        return None
    key = load_key()
    with open(CONFIG_PATH, "rb") as f:
        enc = f.read()
    return decrypt_config(enc, key)
