# Weryfikacja Floteki

Ta instrukcja opisuje sprawdzanie bieżącej wersji. Wyniki każdego uruchomienia
są zapisywane w raportach; dokument nie zastępuje testów konkretnego wydania.

## Testy aplikacji

Wymagane są Windows, Python 3.13 z Tcl/Tk i dostęp do pulpitu:

```powershell
python -m pip install -r requirements-build.txt
python -m pytest tests --capture=sys -q -p no:cacheprovider --junitxml=build/test-results.xml
```

Testy obejmują bazę i historię, szkice, eksport XLSX, interfejs, wersjonowanie,
pakowanie oraz skrypt publikacji. Testy publikacji używają lokalnego zamiennika
GitHub CLI i nie tworzą rzeczywistych wydań. Testy gotowych skoroszytów wymagają
wcześniejszej budowy XLSM.

## Budowa i sprawdzenie EXE

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build_exe.ps1 -Tag V1.0.0.0
```

Skrypt buduje EXE, uruchamia jego autotest na osobnej bazie i sprawdza zgodność
wersji. Katalog wydania zawiera `exe-test.json`, `build-info.json` oraz sumę SHA-256.

## Sprawdzenie Excela

Wymagane są desktopowy Excel dla Windows i biblioteka `pywin32`:

```powershell
python -m pip install pywin32
python tools/build_excel.py --tag V1.0.0.0
$env:FLOTA_EXCEL_FILE = (Resolve-Path dist/Floteka/Floteka.xlsm).Path
$env:FLOTEKA_EXCEL_VERSION_DIR = (Resolve-Path dist/Floteka).Path
python -m pytest tests/test_excel_layout.py tests/test_excel_version.py -q -p no:cacheprovider
python tools/test_excel.py
python tools/test_excel_regressions.py
python tools/test_deadlines_excel.py
python tools/test_optional_excel.py
python tools/test_export_excel.py --exe dist/Floteka/Floteka.exe
```

Sprawdzenia obejmują wersję obu XLSM, układ, zapis i ponowne otwarcie,
archiwizację, szkice, filtry, kolory terminów, dane opcjonalne i eksport gotowego
EXE. Testy makr pracują na kopiach w `build`; raporty powstają w `docs/*.json`.

## Wydania

`tools/build_release.ps1 -Tag V1.0.0.0` wykonuje lokalne testy, budowę i pakowanie.
Podanie `-dev` zachowuje ten sufiks w wersji. W GitHub Actions taki tag wybiera
gałąź `dev`; szczegóły konfiguracji runnera Excel opisuje główny README.

Udany lokalny test nie potwierdza wykonania workflow na GitHub. Dla wydania
sprawdź wynik zadań EXE, Excel i Release oraz zgodność tagu, wersji i commitu
w `build-info.json` i `excel-build.json`.
