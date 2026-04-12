PYTHON := python3
WINE := wine
WINEPREFIX := $(HOME)/.wine
WIN_PYTHON := C:\\users\\till\\AppData\\Local\\Programs\\Python\\Python312\\python.exe
WIN_PYINSTALLER := C:\\users\\till\\AppData\\Local\\Programs\\Python\\Python312\\Scripts\\pyinstaller.exe
SRC_FILE := noten_verwaltung.py
DIST_DIR := dist_windows
BUILD_DIR := build

.PHONY: all exe clean wine-setup

all: exe

# Windows EXE via Wine + Windows Python + PyInstaller
exe: $(DIST_DIR)/Notenverwaltung.exe

$(DIST_DIR)/Notenverwaltung.exe: $(SRC_FILE)
	@rm -rf $(DIST_DIR)
	WINEPREFIX=$(WINEPREFIX) $(WINE) "$(WIN_PYINSTALLER)" \
		--windowed --onefile --noconfirm --clean \
		--name "Notenverwaltung" \
		--distpath $(DIST_DIR) \
		$(SRC_FILE)
	@echo "Fertig: $(DIST_DIR)/Notenverwaltung.exe"

# Windows Python unter Wine einrichten (falls noch nicht vorhanden)
wine-setup:
	@if ! WINEPREFIX=$(WINEPREFIX) $(WINE) "$(WIN_PYTHON)" --version 2>/dev/null | grep -q "Python"; then \
		echo "Installiere Windows Python 3.12 unter Wine..."; \
		cd /tmp && \
		wget -q "https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe" -O python-installer.exe && \
		WINEPREFIX=$(WINEPREFIX) $(WINE) python-installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 && \
		rm -f python-installer.exe && \
		echo "Installiere PyInstaller unter Windows Python..."; \
		WINEPREFIX=$(WINEPREFIX) $(WINE) "$(WIN_PYTHON)" -m pip install pyinstaller; \
		echo "Windows-Python-Umgebung eingerichtet!"; \
	else \
		echo "Windows Python bereits installiert."; \
	fi

clean:
	@echo "Räume Build-Artefakte auf..."
	@rm -rf build dist __pycache__ *.pyc *.spec
	@echo "Fertig - $(DIST_DIR)/Notenverwaltung.exe bleibt erhalten."