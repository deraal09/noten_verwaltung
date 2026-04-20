"""
Constants und Konfiguration für Notenverwaltung
"""

import os
import sys

# Verzeichnis ermitteln
if getattr(sys, 'frozen', False):
    _APP_DIR = os.path.dirname(sys.executable)
else:
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_FILE = os.path.join(_APP_DIR, "noten.ndat")

# Halbjahre
HALBJAHRE = ["1. Halbjahr", "2. Halbjahr"]

# Standard Gewichtung (Prozent für mündliche Unterrichtsleistungen)
DEFAULT_GEWICHTUNG = 60

# Verschlüsselungsparameter
ITERATIONS = 100000

# Notenschlüssel
NOTENSCHLUESSEL = {"IHK": (1, 6), "BG": (0, 15)}

# Standard Notenschlüssel CSV
DEFAULT_NS_CSV = {
    "IHK": "100,1;99,1.1;98,1.1;97,1.2;96,1.2;95,1.3;94,1.3;93,1.4;92,1.4;91,1.5;90,1.6;89,1.7;88,1.8;87,1.9;86,2;85,2;84,2.1;83,2.2;82,2.3;81,2.4;80,2.5;79,2.6;78,2.7;77,2.7;76,2.8;75,2.9;74,2.9;73,3;72,3.1;71,3.1;70,3.2;69,3.3;68,3.3;67,3.4;66,3.5;65,3.6;64,3.6;63,3.7;62,3.7;61,3.8;60,3.9;59,3.9;58,4;57,4;56,4.1;55,4.1;54,4.2;53,4.3;52,4.3;51,4.4;50,4.4;49,4.5;48,4.6;47,4.6;46,4.7;45,4.7;44,4.8;43,4.8;42,4.9;41,4.9;40,5;39,5;38,5;37,5.1;36,5.1;35,5.2;34,5.2;33,5.3;32,5.3;31,5.4;30,5.4;29,5.5;28,5.6;27,5.6;26,5.6;25,5.6;24,5.6;23,5.6;22,5.7;21,5.7;20,5.7;19,5.7;18,5.7;17,5.7;16,5.8;15,5.8;14,5.8;13,5.8;12,5.8;11,5.9;10,5.9;9,5.9;8,5.9;7,5.9;6,5.9;5,6;4,6;3,6;2,6;1,6;0,6",
    "BG": "100,15;99,15;98,15;97,15;96,15;95,15;94,14;93,14;92,14;91,14;90,14;89,13;88,13;87,13;86,13;85,13;84,12;83,12;82,12;81,12;80,12;79,11;78,11;77,11;76,11;75,11;74,10;73,10;72,10;71,10;70,10;69,9;68,9;67,9;66,9;65,9;64,8;63,8;62,8;61,8;60,8;59,7;58,7;57,7;56,7;55,7;54,6;53,6;52,6;51,6;50,6;49,5;48,5;47,5;46,5;45,5;44,4;43,4;42,4;41,4;40,4;39,3;38,3;37,3;36,3;35,3;34,3;33,3;32,2;31,2;30,2;29,2;28,2;27,2;26,1;25,1;24,1;23,1;22,1;21,1;20,1;19,0;18,0;17,0;16,0;15,0;14,0;13,0;12,0;11,0;10,0;9,0;8,0;7,0;6,0;5,0;4,0;3,0;2,0;1,0;0,0",
}

# Notenschlüssel Aliases
_NS_ALIASES = {"Berufsschule": "IHK", "Berufliches Gymnasium": "BG", "IHK": "IHK", "BG": "BG"}

# UI-Meldungen
MSG_WAEHLEN_SJ_KL = "Bitte Schuljahr und Klasse auswählen!"
MSG_WAEHLEN_SJ_KL_FACH = "Bitte Schuljahr, Klasse und Fach auswählen!"
MSG_WAEHLEN_KL_FACH_SK = "Bitte Klasse, Fach und Schülerin auswählen!"
MSG_EXISTIERT_BEREITS = "'{}' existiert bereits!"
MSG_WAEHLEN_ELEMENT = "Bitte ein Element auswählen!"

# Auto-Save Interval in Millisekunden (60 Sekunden)
AUTO_SAVE_MS = 60000


def get_ns_aliases() -> dict:
    """Gibt das Notenschlüssel-Aliases-Dictionary zurück."""
    return dict(_NS_ALIASES)