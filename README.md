# Floteka

**Kierowcy, tiry i przyczepy — terminy rozładunku i pełna historia.**

Floteka udostępnia aplikację Windows z bazą SQLite oraz niezależny skoroszyt
Excel z makrami. Gotowe pliki są publikowane w GitHub Releases. Lokalna budowa
zapisuje je w `dist/Floteka`; repozytorium przechowuje źródła i narzędzia budowy.

Kolory całych wierszy zależą od **daty rozładunku**: przed dzisiaj — czerwony,
dziś lub jutro — pomarańczowy; późniejsze i puste daty mają zwykłe tło.
Dotyczy bieżących wpisów, historii i eksportu XLSX. Excel ma przycisk
**Kolory terminów: WŁ./WYŁ.** w START i DANE; aplikacja ma pole **Kolory terminów**.
W aplikacji kolory odświeżają się co minutę, również po zmianie dnia, bez
naruszania niezapisanego formularza. Excel przelicza kolory podczas otwierania
i przeliczania skoroszytu (w pozostawionym otwartym pliku można nacisnąć F9).

**Aplikacja nie ma limitu czterech wpisów.** Przycisk **+ Dodaj wpis** wykorzystuje
wolne miejsce albo tworzy kolejne. Lista pokazuje tylko wpisy i szkice;
wyczyszczenie usuwa wiersz z widoku i czyści formularz, zachowując historię.
Przywracanie historii umożliwia wybranie
**Nowego miejsca**, również gdy wszystkie wcześniejsze są zajęte.
Excel zachowuje cztery miejsca na kierowcę.

**Tryb ciemny** przełącza kolory GUI i zachowuje wybór po ponownym uruchomieniu.
Zmiana motywu nie kasuje niezapisanych danych. Czerwony i pomarańczowy terminów
mają mocne wypełnienia w obu motywach oraz w skoroszycie i eksporcie.

**Tabela wpisów:** wybór kierowcy i wyszukiwanie po nazwie lub tirze
są nad tabelą. Tabela zajmuje całą szerokość zakładki. **+ Dodaj wpis**,
**Edytuj wpis** albo dwuklik wiersza otwierają kompaktowy formularz pod tabelą.
**Zwiń edytor** oddaje miejsce liście, zachowując wpisane, niezapisane dane.
Formularz jest początkowo zwinięty; pola opcjonalne można dodatkowo rozwinąć.

**Dane opcjonalne:** telefon kierowcy w Rejestrach, trasa i numer zlecenia
w rozwijanej sekcji formularza wpisu. Zakładka **Kontrola tirów** pozwala
uzupełnić przebieg w km, datę przeglądu, ubezpieczenia i serwisu oraz uwagi.
Pokazuje też liczbę tirów z terminem przekroczonym i w ciągu 30 dni.
Żadne z tych pól nie jest wymagane; same dane opcjonalne wpisu tworzą szkic.
Trasa i numer zlecenia są zapisywane także w historii i przywracane razem z wpisem.

W Excelu trasa i numer zlecenia są w **K:L** arkuszy DANE i ARCHIWUM,
telefon w **E** KIEROWCY, a dane tira w **B:F** TIRY. Rejestry są danymi
bieżącymi; zmiany telefonu lub terminów tira nie poprawiają dawnej historii.
Eksport aplikacji pomija puste miejsca i zawiera opcjonalne dane wpisów.
Jeżeli uzupełniono telefon lub dane tirów, eksport bieżących danych dodaje
odpowiednio arkusze KIEROWCY i TIRY. Eksport pojedynczego kierowcy ogranicza
rejestry do jego danych, a eksport wszystkich obejmuje także tiry bez przypisania.

## GitHub Release i wersja z tagu

Workflow `.github/workflows/build-exe.yml` nazywa się **Floteka - Windows EXE**.
Uruchamia się po zmianach kodu na `main`, dla pull requestów do `main`,
dla tagów `v*` i `V*` oraz ręcznie przez **Actions → Floteka - Windows EXE → Run workflow**.
Po ukończeniu wybierz na stronie uruchomienia artefakt **Floteka-Windows-x64**.

Workflow instaluje Python 3.13 x64, uruchamia testy SQLite/eksportu/GUI/nazwy,
buduje `Floteka.exe` i wykonuje test już spakowanego programu. Artefakt zawiera
EXE, instrukcję, sumę SHA-256 i raporty. Nie zawiera baz ani skoroszytów
użytkownika. Budowa EXE nie wymaga własnych sekretów ani zainstalowanego Excela.

Wypchnięcie poprawnego tagu uruchamia także budowę XLSM i publikację
**GitHub Release** z `Floteka.exe`, `Floteka.xlsm`, `Floteka-DEMO.xlsm`,
instrukcją, sumami SHA-256 i raportami. Wydanie jest publikowane dopiero po
pomyślnych testach EXE oraz układu, wersji i makr XLSM.

| Tag | Wersja w plikach | Kod źródłowy | GitHub Release |
| --- | --- | --- | --- |
| `V1.0.0.0` lub `v1.0.0.0` | `1.0.0.0` | Commit wskazany przez tag | Zwykłe wydanie |
| `v1.0.0.0-dev` | `1.0.0.0-dev` | Gałąź `dev` pobrana przy rozpoczęciu budowy | Prerelease, bez oznaczenia Latest |

EXE i XLSM zawsze używają tego samego commitu. Jego SHA jest zapisane w
`build-info.json`, `excel-build.json` i opisie Release. Tag `-dev` nie uruchomi
budowy z innej gałęzi, jeśli `dev` nie istnieje. Zmiana workflow i narzędzi
budowy musi być dostępna również na `dev` przed pierwszym wydaniem testowym.
Warto tworzyć tag testowy na gałęzi `dev`, aby archiwum źródeł generowane przez
GitHub również wskazywało tę gałąź; faktyczny commit binariów podają raporty.

Numer składa się z czterech liczb od 0 do 65535, bez zer wiodących,
z opcjonalnym sufiksem `-dev`. Błędny tag zatrzymuje budowę.
Wersja jest osadzana w tytule aplikacji, `--version`, zasobach Windows EXE,
arkuszu START skoroszytów XLSM oraz właściwościach i stopkach plików Excela.
Eksporty XLSX aplikacji otrzymują wersję programu, który je utworzył.
Budowa bez tagu korzysta z domyślnej wersji zapisanej w `fleet/branding.py`.

### Runner wymagany do budowy XLSM

Zadanie `excel` wymaga własnego runnera GitHub Actions z etykietami
`self-hosted`, `Windows`, `X64`, `excel`. Na tym komputerze zainstaluj
desktopowy Excel i PowerShell 7 (`pwsh`). Uruchom aktualnego runnera
w zalogowanej sesji użytkownika, który może uruchamiać Excel, zamiast jako
usługę bez pulpitu. Dla `actions/setup-python` zainstaluj też 7-Zip i dodaj
`7z` do PATH. Jednorazowe przygotowanie instalacji Pythona w tool cache
wymaga uprawnień administratora; zwykły runner powinien następnie działać
bez podwyższonych uprawnień. Workflow instaluje Python i biblioteki, tworzy nowe, puste
skoroszyty oraz DEMO i testuje makra na kopiach. Nie używa danych użytkownika.
Przygotowanie interpretera opisuje też
[dokumentacja setup-python dla własnych runnerów](https://github.com/actions/setup-python/blob/main/docs/advanced-usage.md#using-setup-python-with-a-self-hosted-runner).
Bez tego runnera zadanie XLSM oczekuje w kolejce, a Release nie zostaje opublikowany.
Dodanie runnera: **Settings → Actions → Runners → New self-hosted runner**;
podczas konfiguracji dodaj etykietę `excel`.

Publikacja korzysta z `GITHUB_TOKEN` i uprawnienia `contents: write` wyłącznie
w zadaniu `release`. Pull requesty i zwykłe zmiany gałęzi nie uruchamiają
runnera Excel ani publikacji. Ponowne uruchomienie może dokończyć szkic Release;
już opublikowane pliki nie są automatycznie zastępowane.

Plik workflow musi najpierw trafić do repozytorium na GitHub. Aby zobaczyć
przycisk ręcznego uruchomienia, umieść go na domyślnej gałęzi `main`.

Samodzielna budowa EXE bez Excela, po instalacji Pythona z Tcl/Tk:

```powershell
python -m pip install -r requirements-build.txt
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build_exe.ps1
```

Wynik: `dist/Floteka/Floteka.exe`. Lokalny pełny pakiet z XLSM wymaga Excela
i pywin32 oraz skryptu `tools/build_release.ps1`. Wersję można podać jawnie:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build_exe.ps1 -Tag v1.0.0.0-dev
python tools/build_excel.py --tag v1.0.0.0-dev
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build_release.ps1 -Tag V1.0.0.0
```

Lokalny parametr `-Tag` / `--tag` ustawia wersję aktualnego checkoutu;
wybór gałęzi `dev` odbywa się w GitHub Actions.

## Który plik uruchomić

<<<<<<< HEAD
| Plik | Zastosowanie |
| --- | --- |
| `Floteka.xlsm` | Pusty skoroszyt do właściwej pracy w Excelu. |
| `Floteka-DEMO.xlsm` | Oddzielny skoroszyt z fikcyjnymi przykładami. |
| `Floteka.exe` | Samodzielna aplikacja Windows 64-bit z SQLite. Python nie jest potrzebny. |
| `Uruchom-DEMO.bat` | Uruchamia EXE na osobnej przykładowej bazie `Floteka-DEMO.sqlite3`. |
| `Zrodla.zip` | Kod aplikacji, VBA, testy i narzędzia budowy. |
=======
| Plik                        | Zastosowanie                                                              |
| --------------------------- | ------------------------------------------------------------------------- |
| `Floteka.xlsm`            | Pusty skoroszyt do właściwej pracy w Excelu.                            |
| `Floteka-DEMO.xlsm`       | Oddzielny skoroszyt z fikcyjnymi przykładami.                            |
| `Floteka-TWOJE-DANE.xlsm` | Poprawiona kopia dotychczasowych danych, jeśli dołączona do pakietu.   |
| `Floteka.exe`             | Samodzielna aplikacja Windows 64-bit z SQLite. Python nie jest potrzebny. |
| `Uruchom-DEMO.bat`        | Uruchamia EXE na osobnej przykładowej bazie `Floteka-DEMO.sqlite3`.    |
| `Zrodla.zip`              | Kod aplikacji, VBA, testy i narzędzia budowy.                            |
>>>>>>> 97187ab0e2c734588af47313a1871edbf670bf46

Jeśli pobierasz pakiet ZIP, rozpakuj go do wybranego folderu. Skoroszyty i aplikacja
nie synchronizują danych. Eksport aplikacji tworzy zwykły plik XLSX z danymi.

## Zasady wspólne

- Kolumny A:F: kierowca, rejestracja tira, rejestracja przyczepy,
  informacje dodatkowe, data załadunku, data rozładunku.
- Wpis jest kompletny, gdy ma kierowcę, tira, przyczepę i obie poprawne daty.
  Informacje dodatkowe są opcjonalne. Rozładunek nie może poprzedzać załadunku.
- Niekompletny wpis to **szkic**. Zostaje w swoim miejscu, ale nie jest
  kopiowany do historii. Przy jego czyszczeniu pojawia się potwierdzenie.
- Zajęte miejsca pozostają na liście; aplikacja pozwala dopisać kolejne.
  Nic nie jest automatycznie usuwane z bieżących wpisów.
- **Zwolnij najstarsze** czyści miejsce zajęte najwcześniej. Liczy się kolejność
  zajęcia miejsca, a nie data załadunku. Późniejsza edycja nie zmienia kolejności.
- Przycisk **× / x** przy wierszu czyści tylko wskazane miejsce. Pozostałe
  miejsca nie przesuwają się. Kierowca i tir pozostają.
- Historia nie jest usuwana przez te operacje. Poprawka kompletnego wpisu
  dopisuje nową wersję. Identyczne dane dla tego samego
  kierowcy nie jest dopisywana ponownie, także po przywróceniu do innego miejsca.
- Jeżeli dwa różne zlecenia mają identyczne dane, rozróżnij je numerem
  w opcjonalnym polu „Numer zlecenia”.
- Przywracanie historii wymaga pustego miejsca lub utworzenia nowego w aplikacji. Przywrócone dane używają
  **obecnego tira** kierowcy. Przy zmianie tira pojawi się informacja;
  historyczny numer pozostaje w starej kopii.
- Zmiana nazwy kierowcy lub jego tira zachowuje wcześniejsze kopie historii.
  Kierowców o tej samej nazwie rozróżniają stałe identyfikatory.

## Excel — rozpoczęcie pracy

Wymagany jest klasyczny Microsoft Excel dla Windows z obsługą VBA.
Wersja przeglądarkowa nie uruchamia przycisków VBA.

1. Otwórz `Floteka.xlsm`. Jeśli Excel wyświetli pasek o zawartości,
   zezwól na makra w tej sprawdzonej kopii pliku.
2. W START lub DANE kliknij **+ Dodaj kierowcę**. Podaj nazwę i rejestrację
   tira. Powstaną cztery wiersze w DANE; nowy tir zostanie dopisany do rejestru.
3. Edytuj kolumny **C:F**. Przyczepę możesz wybrać z listy albo wpisać nowy
   numer. Nowy numer zostanie dodany do rejestru przyczep. Daty: `17.11.2026`.
4. Zapisuj skoroszyt przez **Ctrl+S** lub **Zapisz wszystko**. Zmiany i historia
   są trwałe dopiero po zapisaniu pliku.

W DANE zaznacz dowolny wiersz kierowcy przed kliknięciem **Zwolnij najstarsze**.
Przycisk **x** po prawej stronie działa na swój wiersz.

W KIEROWCY możesz poprawić nazwę, przypisać innego tira lub wpisać **NIE**
w kolumnie Aktywny. Wyłączony kierowca zostaje ukryty w DANE; jego wpisy
i historia pozostają. Ustaw **TAK**, aby znów pokazać jego cztery miejsca.
Nowego kierowcę można też dopisać w pustym wierszu KIEROWCY, uzupełniając nazwę
i tira. Kolumna ID jest obsługiwana automatycznie.

W TIRY i PRZYCZEPY dopisuj rejestracje poniżej dotychczasowych. Rejestry mają
przygotowane 5000 edytowalnych pozycji. DANE powiększa się przez dodawanie kierowców.

W ARCHIWUM używaj filtrów w nagłówkach. Aby przywrócić wpis, zaznacz jego wiersz,
kliknij **Przywróć wpis** i podaj numer pustego miejsca od 1 do 4.
W START jest przycisk **Kopia pliku**, zapisujący oddzielny XLSM.
Filtry służą do przeglądania i nie przenoszą wpisów do historii.
Używaj filtrowania bez zdejmowania ochrony arkusza i zmiany kolejności miejsc.

Arkusze mają ochronę przed przypadkowymi zmianami układu i historii.
Nie jest to zabezpieczenie przed celową ingerencją w plik. Kod VBA znajduje się
w źródłach w folderze `excel`. Do używania skoroszytu nie trzeba włączać
programistycznego dostępu do projektu VBA w ustawieniach Excela.

**Bez makr:** oglądanie danych jest możliwe, lecz przyciski, walidacja VBA
i zapis historii nie działają. Pracuj z włączonymi makrami; zachowaj rozszerzenie XLSM.
Nie zapisuj skoroszytu jako XLSX, ponieważ spowoduje to usunięcie makr.

## Aplikacja Windows — rozpoczęcie pracy

1. Uruchom `Floteka.exe`.
2. W zakładce **Rejestry** dodaj tira, a potem kierowcę i jego przypisanie.
   Przyczepy możesz dodać wcześniej albo wpisać podczas edycji miejsca.
3. W **Bieżących wpisach** wybierz kierowcę i miejsce lub kliknij **+ Dodaj wpis**. Uzupełnij formularz,
   następnie kliknij **Zapisz szkic / wpis** lub naciśnij **Ctrl+S**.
4. Otwórz **Historię**, aby filtrować po kierowcy, wyszukiwać tekst i przywracać
   wcześniejsze wpisy do pustego miejsca.

Daty można wpisać jako `17.11.2026`, `17/11/2026` lub `2026-11-17`.
Po zapisaniu formularza transakcja trafia od razu na dysk. Przed zmianą
kierowcy, miejsca lub zamknięciem programu pojawia się pytanie o niezapisane
zmiany. Błąd formularza pozostawia wprowadzone wartości do poprawienia.

**Zmień nazwę / przypisz tira** w Rejestrach edytuje kierowcę i jego jeden
bieżący tir. **Włącz / wyłącz z listy** pozwala ukryć kierowcę lub pojazd.
Wyłączenie kierowcy nie kasuje także jego szkiców; wrócą po ponownym włączeniu.
Tira przypisanego aktywnemu kierowcy trzeba najpierw zastąpić innym.

### Baza i kopie

Zwykłe uruchomienie używa bazy:

```text
%LOCALAPPDATA%\Floteka\flota.sqlite3
```

Przycisk **Folder danych** na dole okna otwiera folder aktualnej bazy.
Przeniesienie EXE do innego folderu nie zmienia tej bazy.

Przycisk **Kopia bazy** tworzy kompletną, spójną kopię SQLite, także podczas
działania programu. Baza zawiera szkice, rejestry i całą historię. Triggery
blokują zwykłe polecenia UPDATE i DELETE tabeli historii. Nie usuwaj ręcznie
plików bazy; kopie zapasowe przechowuj dodatkowo poza tym komputerem.

Aby otworzyć kopię lub własną bazę, zamknij program i uruchom:

```powershell
.\Floteka.exe --db "D:\Kopie\Flota-kopia.sqlite3"
```

Możesz przygotować skrót Windows z takim argumentem. Ścieżka do własnej bazy
nie zmienia bazy domyślnej. Pracuj na lokalnym dysku jednego stanowiska;
narzędzie nie jest przeznaczone do współdzielonego pliku na dysku sieciowym.

### Eksport

**Eksport do Excela** pozwala wybrać bieżące dane, pełną historię lub oba
arkusze oraz wszystkich albo jednego kierowcę. Dane bieżące wszystkich
kierowców obejmują aktywnych; historia wszystkich obejmuje także wyłączonych.
Eksport pojedynczego kierowcy może obejmować również kierowcę wyłączonego.

Daty są datami Excela, a tekst pozostaje tekstem, także gdy zaczyna się od `=`.
XLSX służy do przeglądania i dalszej analizy. Zmiana tego pliku nie aktualizuje
bazy ani skoroszytu XLSM. Tekst przekraczający limit 32767 znaków komórki
Excela zatrzymuje eksport z komunikatem; baza zachowuje pełną treść.

## Uruchamianie i budowanie ze źródeł

Źródła wymagają Pythona 3.13 z Tcl/Tk. W folderze źródeł:

```powershell
python -m pip install -r requirements.txt
python main.py
```

`Uruchom.bat` uruchamia tę samą wersję ze źródeł. Można użyć argumentu `--db`.
Testy: `python -m pytest tests --capture=sys -q` (narzędzia z `requirements-build.txt`
i dostęp do pulpitu dla testów Tk; `--capture=sys` omija konflikt uchwytów Tcl/Tk na Windows).
Test układu domyślnie sprawdza `dist/Floteka/Floteka.xlsm`, jeśli istnieje.
Własny plik można wskazać zmienną
`FLOTA_EXCEL_FILE`; używają jej również testy makr. Pomiar i testy filtrów:
`python tools/check_excel_filters_speed.py dist/Floteka/Floteka.xlsm`.
Budowa EXE i XLSM: `tools/build_release.ps1`; wymagane są PyInstaller,
pywin32 i zainstalowany Excel dla Windows.

Budowa XLSM chwilowo włącza AccessVBOM w profilu użytkownika i przywraca
poprzednią wartość w bloku `finally`. Nie zmienia globalnego ustawienia
zezwalającego na wykonywanie wszystkich makr. Ten krok jest potrzebny
wyłącznie do ponownego zbudowania pliku ze źródeł.

Test EXE: `Floteka.exe --self-test raport.json` używa osobnej tymczasowej
bazy. Testy Excela w `tools/test_excel*.py` pracują wyłącznie na kopiach w `build`.
Procedura sprawdzania wydań znajduje się w [docs/verification.md](docs/verification.md),
dołączanym do lokalnego pakietu jako `WERYFIKACJA.md`. Wyniki konkretnego wydania
są zapisane w raportach dołączonych do artefaktów.
