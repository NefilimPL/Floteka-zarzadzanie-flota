# Floteka

**Kierowcy, tiry i przyczepy — terminy rozładunku i pełna historia.**

Aktualne wydanie znajduje się w `dist/Floteka`. Wyniki budowania EXE/XLSM/ZIP
są pomijane przez Git; repozytorium przechowuje źródła i narzędzia budowy.
Starsze lokalne dostawy zachowano w `local-archive`, poza drzewem źródeł,

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
**Nowego miejsca**, również gdy wszystkie wcześniejsze są zajęte. Dotychczasowa
baza aktualizuje się automatycznie przy uruchomieniu, zachowując wpisy i historię.
Po aktualizacji baza wymaga nowej wersji aplikacji. Excel zachowuje cztery miejsca.

**Tryb ciemny** przełącza kolory GUI i zachowuje wybór po ponownym uruchomieniu.
Zmiana motywu nie kasuje niezapisanych danych. Czerwony i pomarańczowy terminów
mają mocne wypełnienia w obu motywach oraz w skoroszycie i eksporcie.

**Większa tabela wpisów:** wybór kierowcy i wyszukiwanie po nazwie lub tirze
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

## Automatyczna budowa EXE w GitHub Actions

Workflow `.github/workflows/build-exe.yml` nazywa się **Floteka - Windows EXE**.
Uruchamia się po zmianach kodu na `main`, dla pull requestów do `main`,
dla tagów `v*` oraz ręcznie przez **Actions → Floteka - Windows EXE → Run workflow**.
Po ukończeniu wybierz na stronie uruchomienia artefakt **Floteka-Windows-x64**.

Workflow instaluje Python 3.13 x64, uruchamia testy SQLite/eksportu/GUI/nazwy,
buduje `Floteka.exe` i wykonuje test już spakowanego programu. Artefakt zawiera
EXE, instrukcję, sumę SHA-256 i raporty. Nie zawiera baz ani skoroszytów
użytkownika. Nie wymaga własnych sekretów ani zainstalowanego Excela.

Plik workflow musi najpierw trafić do repozytorium na GitHub. Aby zobaczyć
przycisk ręcznego uruchomienia, umieść go na domyślnej gałęzi `main`.

Samodzielna budowa EXE bez Excela, po instalacji Pythona z Tcl/Tk:

```powershell
python -m pip install -r requirements-build.txt
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build_exe.ps1
```

Wynik: `dist/Floteka/Floteka.exe`. Pełny pakiet z XLSM wymaga lokalnego Excela
i pywin32 oraz skryptu `tools/build_release.ps1`. GitHub Actions buduje aplikację
SQLite; testy VBA i budowa skoroszytu pozostają częścią lokalnego procesu.

Domyślna baza znajduje się w `%LOCALAPPDATA%\Floteka\flota.sqlite3`.
Przy pierwszym uruchomieniu nowej wersji istniejąca baza ze starego folderu
`FlotaAdriana` jest kopiowana przez mechanizm kopii SQLite, również z danymi
w dzienniku WAL. Nowa baza nigdy nie jest nadpisywana, a stary plik pozostaje
kopią bezpieczeństwa. Własna ścieżka `--db` pomija tę migrację.

Poprawiona wersja usuwa nakładanie przycisków na opis, przywraca nagłówki
w odpowiednich kolumnach i naprawia dopisywanie historii przy aktywnych filtrach.
Jeżeli otrzymałeś `Floteka-TWOJE-DANE.xlsm`, jest to poprawiona kopia
dotychczasowego skoroszytu: dane, szkice i historia zostały porównane z oryginałem.
Oryginał nie został zmieniony. Do dalszej pracy używaj poprawionej kopii.

Filtry nie przenoszą wpisów do historii. Indeks duplikatów jest odtwarzany
z rzeczywistych rekordów, również ukrytych i odfiltrowanych. Dopisywanie nie
usuwa kryteriów filtra. Ochrona arkusza blokuje sortowanie, które zmieniłoby
kolejność czterech miejsc; używaj filtrowania bez zdejmowania ochrony.

Pomiar na tym komputerze: 250 kierowców, 1000 pełnych wpisów i 1000 kopii historii.
Mediana zapisu niezmienionego pliku w aktualnej próbie wyniosła około 0,33 s.
To wynik tej próby, nie gwarancja czasu dla dowolnej wielkości pliku i komputera.

Dwa niezależne narzędzia do zarządzania kierowcami, tirami i przyczepami.
Każdy kierowca ma jednego bieżącego tira. Excel ma **cztery miejsca na wpisy**,
a aplikacja pokazuje tylko zajęte wpisy i pozwala dodawać następne bez stałego limitu.
Kompletne wpisy są od razu kopiowane do historii, nawet gdy nadal są widoczne.

## Który plik uruchomić

| Plik                        | Zastosowanie                                                              |
| --------------------------- | ------------------------------------------------------------------------- |
| `Floteka.xlsm`            | Pusty skoroszyt do właściwej pracy w Excelu.                            |
| `Floteka-DEMO.xlsm`       | Oddzielny skoroszyt z fikcyjnymi przykładami.                            |
| `Floteka-TWOJE-DANE.xlsm` | Poprawiona kopia dotychczasowych danych, jeśli dołączona do pakietu.   |
| `Floteka.exe`             | Samodzielna aplikacja Windows 64-bit z SQLite. Python nie jest potrzebny. |
| `Uruchom-DEMO.bat`        | Uruchamia EXE na osobnej przykładowej bazie `Floteka-DEMO.sqlite3`.    |
| `Zrodla.zip`              | Kod aplikacji, VBA, testy i narzędzia budowy.                            |

Najpierw rozpakuj cały pakiet ZIP do wybranego folderu. Skoroszyty i aplikacja
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
  dopisuje nową wersję. Identyczna zawartość sześciu pól dla tego samego
  kierowcy nie jest dopisywana ponownie, także po przywróceniu do innego miejsca.
- Jeżeli dwa rzeczywiście różne zlecenia mają identyczne dane, rozróżnij je
  numerem zlecenia w „Informacjach dodatkowych”.
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
Wersja 1.1.2 naprawia konflikt filtrów arkusza i tabeli powodujący komunikat
Excela o naprawianiu pliku. Eksporty utworzone wcześniejszą wersją wygeneruj
ponownie z aplikacji. Test `tools/test_export_excel.py --exe dist/Floteka/Floteka.exe`
otwiera XLSX z gotowego programu w rzeczywistym Excelu bez trybu naprawy.
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
Testy: `python -m pytest -q` (wymaga pytest i dostępu do pulpitu dla testów Tk).
Test układu domyślnie sprawdza `dist/Floteka/Floteka.xlsm`, jeśli istnieje,
a dla starszych dostaw `dist/poprawione/Flota-Adriana.xlsm` lub
`dist/Flota-Adriana.xlsm`. Własny plik można wskazać zmienną
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
Wyniki wykonanych sprawdzeń znajdują się w dołączonym `WERYFIKACJA.md`.
