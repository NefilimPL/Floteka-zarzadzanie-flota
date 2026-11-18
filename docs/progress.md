# Ledger — plan: docs/superpowers/plans/2026-11-17-flota.md

Zatwierdzony projekt: docs/superpowers/specs/2026-11-17-zarzadzanie-flota-design.md.
Wykonanie w bieżącym, pustym repozytorium; użytkownik zatwierdził wykonanie obu produktów.
Task 1 -> Task 2: GUI używa interfejsu Store i export_xlsx z planu; zgodne.
Task 1 -> Task 3: wspólne zasady historii, niezależne przechowywanie; zgodne.
Task 2/3 -> Task 4: źródła oraz gotowe artefakty wymagają osobnego sprawdzenia; zgodne.

Task 1: complete — 17 testów SQLite/eksportu przeszło (python -m pytest tests/test_store.py tests/test_export.py -q).
Task 2: complete — pełne 20 testów przeszło, w tym rzeczywiste Tk GUI (poza piaskownicą, -p no:cacheprovider --basetemp .test-external-01).
Środowisko: Tcl/Tk i Excel COM wymagają procesu poza piaskownicą. Katalogów pytest utworzonych wewnątrz piaskownicy nie używać jako basetemp procesu zewnętrznego.
Excel: budowa przywraca wcześniejszą wartość AccessVBOM; format A4 pomijany, gdy lokalny Excel nie ma drukarki.

Task 3: complete — XLSM zbudowany w lokalnym Excelu; tools/test_excel.py przeszedł 9 grup sprawdzeń; trzy regresje Excel przeszły.
Final review: niezależny code_review wskazał 4 istotne usterki; wszystkie naprawione.
Final: fixed czyszczenie szkicu pustym formularzem — test_saving_empty_fields_over_draft_needs_confirmation i test_empty_save_cancel_preserves_draft RED -> GREEN.
Final: fixed ciche obcinanie eksportu — test_export_refuses_truncation_and_keeps_previous_file RED -> GREEN.
Final: fixed przerwane tworzenie kierowcy — failed_driver_creation w tools/test_excel_regressions.py RED -> GREEN.
Final: fixed wklejanie kilku kierowców z błędem pierwszego wiersza — multirow_registry_paste RED -> GREEN.
Python suite po poprawkach: 24 passed (-p no:cacheprovider --basetemp .test-external-03).

Task 4: complete — tools/build_release.ps1 zakończył się kodem 0: 24 testy, budowa XLSM/EXE, oba scenariusze Excel, świeży test spakowanego EXE oraz ZIP.
Po uporządkowaniu widocznych zakładek XLSM ponownie przeszły test główny i 3 regresje; pakiet przepakowany.
Kontrola ZIP: 8 oczekiwanych plików, bez błędów CRC; źródła GUI i VBA w Zrodla.zip zgodne z aktualnym dyskiem.
Podgląd GUI sprawdzony: wszystkie cztery miejsca, formularz, przyciski oraz komunikaty i ścieżka bazy widoczne.

Poprawki po zgłoszeniu zasłoniętego tekstu i pytaniu o filtry:
- RED -> GREEN: gotowy XLSM miał nagłówki tylko w H5/I5 (Python COM Resize wybierał komórkę zamiast zakresu); zastąpiono jawnym Range. Przyciski przeniesione do własnego wiersza 3, kotwice OOXML sprawdzone.
- Odtworzono błąd End(xlUp) przy filtrze: LastRow pomijało ukrytą historię, możliwe nadpisania i duplikaty. Naprawa odczytuje fizyczne Value2 do UsedRange; nie polega na End ani Find.
- LoadKeys odtwarza indeks z A:F i ID kierowcy w H, także dla starych brakujących/błędnych kluczy J. Nowy indeks J jest zapisywany i odczytywany jawnie.
- RefreshArchiveFilter zachowuje kryteria. Przeszły brak widocznych wyników, wielokrotny wybór i zakres dat; każde pierwotne pole 1000 rekordów pozostało identyczne.
- AddPlate korzysta z unieważnianego cache. FlushAll porównuje tablice, kontroluje obecność historii i przetwarza tylko zmienione lub brakujące kopie.
- Pomiar 250 kierowców/1000 wpisów/1000 historii: mediana Save 6.8059 s -> 0.2089 s.
- Przegląd review_excel_fixes: uwagi o rekonsyliacji historii i domyślnej ścieżce testów naprawione; regresje napisane i wykonane RED -> GREEN.
- Nowe pliki są w dist/poprawione. Zmieniony oryginał użytkownika zachowany bez zmiany SHA-256; kopia TWOJE-DANE zachowuje wszystkie komórki danych/rejestrów/_STAN/historii.

Floteka + GitHub Actions:
- Zmieniono nazwę publiczną, CLI --version, właściwości EXE, pliki XLSM i ZIP. Zachowano %LOCALAPPDATA%/FlotaAdriana/flota.sqlite3.
- tests/test_branding.py i tytuł rzeczywistego GUI: RED -> GREEN. Pełna lokalna seria: 28 passed.
- .github/workflows/build-exe.yml: windows-2025/Python3.13 x64, push main/tags, PR main, workflow_dispatch. Testy bez Excela, build_exe.ps1, jawne artefakty EXE/reporty/instrukcja/SHA bez baz i skoroszytów.
- build_exe.ps1 lokalnie zbudował i uruchomił gotowy Floteka.exe; raport frozen/app/version poprawny i ProductName=Floteka potwierdzony z zasobu Windows.
- YAML zweryfikowany parserem PyYAML poza piaskownicą. GitHub Actions NIE został wypchnięty ani uruchomiony na serwerze.
- Przegląd review_floteka_ci: brak istotnych uwag; drobną niespójność ścieżek w README poprawiono.
- build_release.ps1 przeszedł testy aplikacji, układu, makr, 5 regresji i filtrów oraz zbudował pełny pakiet.
- Migracja poprawionego XLSM jest idempotentna dla naszych przycisków; ignoruje wirtualne kontrolki filtrów z kolekcji Shapes. Floteka-TWOJE-DANE zachowała każdą komórkę danych i oryginalny hash źródła.
