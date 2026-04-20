"""
Verschlüsselungsfunktionen für Notenverwaltung
"""

import hashlib
import os
from typing import Optional, Dict, Any

from constants import ITERATIONS

logger = __import__('logging').getLogger(__name__)


def _derive_key(password: str, salt: bytes) -> bytes:
    """Leitet einen Schlüssel aus dem Passwort und Salt ab."""
    return hashlib.pbkdf2_hmac('sha256', password.encode(), salt, ITERATIONS)


def _xor_encrypt(data_bytes: bytes, key: bytes) -> bytes:
    """XOR-Verschlüsselung mit periodischem Schlüssel."""
    key_len = len(key)
    return bytes(b ^ key[i % key_len] for i, b in enumerate(data_bytes))


def encrypt_data(data_dict: Dict[str, Any], password: str) -> bytes:
    """
    Verschlüsselt ein Dictionary in ein Byte-Array.

    Format: salt (16 bytes) + verify hash (32 bytes) + encrypted data
    """
    import json
    import zlib
    import base64

    json_bytes = json.dumps(data_dict, ensure_ascii=False).encode('utf-8')
    compressed = zlib.compress(json_bytes)
    salt = os.urandom(16)
    key = _derive_key(password, salt)
    encrypted = _xor_encrypt(compressed, key)
    verify = hashlib.sha256(key).digest()
    return base64.b64encode(salt + verify + encrypted)


def decrypt_data(raw_bytes: bytes, password: str) -> Optional[Dict[str, Any]]:
    """
    Entschlüsselt ein Byte-Array in ein Dictionary.

    Returns None bei Fehlern (falsches Passwort, beschädigte Daten).
    """
    import json
    import zlib
    import base64

    try:
        raw = base64.b64decode(raw_bytes)
    except Exception:
        logger.warning("Base64-Dekodierung fehlgeschlagen")
        return None

    if len(raw) < 48:
        logger.warning("Daten zu kurz für Entschlüsselung")
        return None

    salt = raw[:16]
    verify = raw[16:48]
    encrypted = raw[48:]

    key = _derive_key(password, salt)
    if hashlib.sha256(key).digest() != verify:
        logger.warning("Passwort-Verifikation fehlgeschlagen")
        return None

    compressed = _xor_encrypt(encrypted, key)

    try:
        json_bytes = zlib.decompress(compressed)
        return json.loads(json_bytes.decode('utf-8'))
    except (zlib.error, json.JSONDecodeError, UnicodeDecodeError) as e:
        logger.error("Dekompression/Deserialisierung fehlgeschlagen: %s", e)
        return None