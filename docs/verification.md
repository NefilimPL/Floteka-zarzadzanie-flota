# Weryfikacja Floteki

Ta instrukcja opisuje sprawdzanie bieżącej wersji. Wyniki każdego uruchomienia
są zapisywane w raportach; dokument nie zastępuje testów konkretnego wydania.

## Lokalny pakiet bez Office

Na Windows z Pythonem 3.13 x64, Tcl/Tk, pip i dostępem do pulpitu uruchom
`Buduj.bat`. Pierwsze pobranie zależności wymaga internetu. Skrypt instaluje je
w `build/local-venv`, tworzy oba XLSM ze sprawdzonych szablonów, wykonuje pełne
testy aplikacji oraz struktury, układu i wersji XLSM, buduje i testuje EXE,
a następnie pakuje pliki. Każdy etap musi zakończyć się powodzeniem.

Wynik trafia do nowego `dist/Floteka-lokalnie/<data-godzina>`; własny katalog
podany przez `Buduj.bat -OutputDirectory "D:\Pakiety\Nowy"` musi być pusty.
Wersję ustawia `-Tag V1.0.0.0` lub `-Tag v1.0.0.0-dev`. Błąd zatrzymuje przebieg;
zapis konsoli znajduje się w `build.log` w katalogu wyniku.

Sprawdź raport `test-results.xml`, `exe-test.json` (`ok`, `frozen`, nazwa i wersja),
`build-info.json` (wersja, tag, commit i SHA-256 EXE) oraz `excel-build.json`
(zgodna wersja, tag, commit i SHA-256 obu XLSM). Pakowanie ponownie sprawdza
zgodność raportów z rzeczywistymi plikami. `SUMY-SHA256.txt` w katalogu wyniku
obejmuje również ZIP pakietu; lista wewnątrz ZIP obejmuje jego zawartość.
`Zrodla.zip` zawiera launcher, potrzebne narzędzia i szablony. W paczce nie ma
użytkowych baz, migrowanych skoroszytów ani środowiska budowy.

Ten przebieg nie uruchamia Excela: `excel-build.json` jawnie zapisuje
`engine: "template"` i `macro_tests_run: false`. Zachowanie binarnego projektu
VBA i przycisków oraz testy struktury nie potwierdzają wykonania makr podczas
tej budowy. Do ponownego sprawdzenia makr służy osobna procedura poniżej.
Aktualizacja VBA lub układu wymaga `python -m tools.refresh_excel_templates`
na komputerze z Excelem i pywin32 oraz ponownych testów makr. Niezgodność sum
szablonów albo powiązanych źródeł blokuje zwykłą budowę.

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

Udany lokalny test nie potwierdza wykonania workflow na GitHub. Dla domyślnego
wydania bez Office sprawdź wynik zadań EXE i Release, pominięcie zadania Excel
oraz tag, wersję i commit w `build-info.json`. Budowę XLSM włącza zmienna
repozytorium `FLOTEKA_BUILD_EXCEL=true`; wymaga to runnera z Microsoft Excelem.
W tym trybie sprawdź również zadanie Excel i zgodność obu raportów
`build-info.json` i `excel-build.json`. Błąd włączonej budowy XLSM musi blokować
publikację zamiast tworzyć niepełne wydanie.
