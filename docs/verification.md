# Weryfikacja — Floteka 1.1.2

Aktualizacja eksportu XLSX z 18.11.2026 dla Windows 64-bit. Skoroszyty XLSM
i sprawdzenia makr zachowane z wydania 1.1.0 z 17.11.2026. Eksport XLSX
sprawdzany oddzielnie w rzeczywistym Excelu 16.0; VBA i baza bez zmian.

## Wykonane sprawdzenia

- **78 testów Python/SQLite/eksportu/GUI/układu XLSM/nazwy** — wszystkie przeszły.
- **Naprawa eksportu XLSX** — odtworzono błąd normalnego otwarcia w Excelu.
  Przyczyną był filtr arkusza nałożony na zakres tabeli mającej własny filtr.
  Usunięto drugi filtr; tabele zachowują swoje AutoFilter, a arkusze bez tabel
  zachowują filtr arkusza. Sprawdzono zakresy bieżący, historia, oba naraz,
  opcjonalne rejestry, pustą bazę i szkic. Testy sprawdzają wszystkie wartości
  komórek, brak formuł użytkownika, tabele i filtry, także po SaveCopyAs i
  ponownym otwarciu. XLSX z gotowego EXE otwierany jest z CorruptLoad=0,
  bez wywołania odzyskiwania. Raport: `docs/excel-export.json`.
- **Izolacja autotestu** — domyślny autotest EXE trzyma eksport w swoim
  tymczasowym katalogu. Osobny plik powstaje wyłącznie po podaniu
  `--self-test-export`; raport i eksport muszą być różnymi plikami. Testy
  potwierdzają zachowanie istniejących skoroszytów obok raportu.
- **Większa tabela** — w oknie 1366×700 rzeczywisty Tk pokazał 13 pełnych
  wierszy po zwinięciu edytora i 8 pełnych wierszy przy rozwiniętym formularzu
  z polami opcjonalnymi. Testy sprawdziły 1120×700, 1366×700 oraz 1600×748,
  wybór i wyszukiwanie kierowcy, anulowanie zmiany i zachowanie niezapisanego
  formularza po zwinięciu. Tabela zajmuje co najmniej 90% szerokości okna.
  Przy skali 200%, ustawionej przed budową kontrolek, paski akcji zawijają
  się do dwóch rzędów. Testy 1040×700 i 1120×700 potwierdziły pełne granice
  i rozmiary widocznych przycisków, w tym czyszczenia i zapisu.
- **Folder danych** — domyślna ścieżka ma nazwę Floteka. Test przejęcia starej
  bazy obejmuje nieprzeniesione z WAL dane, brak nadpisania istniejącej nowej
  bazy i bezpieczną odmowę przy uszkodzonym źródle. Nie uruchamiano migracji
  na używanej bazie użytkownika; wykona ją nowa aplikacja przy pierwszym starcie.
- **Puste wpisy i motyw** — GUI ukrywa puste miejsca i wykorzystuje je ponownie.
  Usunięcie wyczyściło formularz, pozostawiając historię. Ciemny motyw jest
  zapisywany w bazie; przełączenie zachowało niezapisane pola opcjonalne.
  Podglądy: `docs/podglad-jasny.png`, `docs/podglad-ciemny.png`
  i `docs/podglad-kontrola-tirow.png`.
- **Opcjonalne informacje** — telefon, trasa, numer zlecenia, przebieg,
  przegląd, ubezpieczenie, serwis i uwagi tira. Test eksportu ponownie
  odczytał XLSX, sprawdzając daty, telefon, bezpieczny tekst zaczynający
  się od znaku równości i pomijanie pustych miejsc.
- **Sześć scenariuszy opcjonalnych w Excelu 16.0** — szkic, potwierdzenie
  czyszczenia, wersjonowanie i przywracanie K:L, filtr trasy, walidacja,
  wyraźne kolory, restart bez duplikatów oraz zapis rejestrów bez makr
  i późniejsza błędna edycja bez utraty wcześniejszych danych.
- **Poprawki po przeglądzie** — cztery odtworzone regresje: anulowanie wyboru
  przy niewidocznym nowym wpisie, zapis szkiców z dwóch okien, równoczesna
  migracja starej bazy i nieaktualna kopia rejestrów po zapisie bez makr.
  Przypadki dostały testy, które najpierw wykazały błąd, a po poprawkach przeszły.
- **Terminy rozładunku** — rzeczywisty Excel 16.0 potwierdził kolor całego
  wiersza: wczoraj czerwony, dziś/jutro pomarańczowy, pojutrze i brak daty
  zwykłe tło. Reguły działają w DANE, ARCHIWUM i w nowych blokach kierowców.
  Wyłączenie kolorów zachowuje się po ponownym otwarciu XLSM. Eksport XLSX
  zawiera dynamiczne reguły TODAY(), sprawdzone przez ponowne odczytanie pliku.
- **Dowolna liczba miejsc w aplikacji** — zapisano do 12 miejsc i przywrócono
  historię do trzynastego; restart zachował wpisy. Test starego schematu SQLite
  v1 i v2 potwierdziły migrację do v3, zachowanie szkiców, historii i kolejności
  zajęcia oraz integralność kluczy obcych. Nieprawidłowe miejsca nie tworzą historii.
- **Przycisk kierowcy** — test odtworzył niewidoczny przycisk, a po zmianie
  układu potwierdził pełną widoczność przy 1120×780 i dwóch skalach Tk.
  Przełączanie kolorów nie usuwa niezapisanego formularza. Podgląd okna
  po zmianach zapisano w podglądach jasnego i ciemnego motywu.
- **Porządki** — wygenerowane pliki dist są pomijane przez Git. Starsze lokalne
  wydania oraz oryginał skoroszytu użytkownika zachowano w `local-archive`.
  Skoroszytu użytkownika nie było lokalnie: odtworzono egzemplarz z poprzedniego
  wydania w commicie c8b2ab0 do `local-archive/skoroszyt-z-historii-c8b2ab0`.
  Zaktualizowaną kopię `Floteka-TWOJE-DANE.xlsm` porównano z tym źródłem,
  zachowując dane, szkice, historię i sumę SHA-256 oryginału. Nie potwierdzano
  zgodności z ewentualnymi nowszymi danymi poza repozytorium.
- **GitHub Actions** — lokalnie wykonano używany przez workflow skrypt
  `tools/build_exe.ps1`, w tym uruchomienie gotowego Floteka.exe.
  Workflow nie był wypchnięty ani uruchomiony na serwerze GitHub w tej sesji.
- **Zmiana nazwy** — GUI, CLI, właściwości EXE, pliki XLSM i pakiet noszą nazwę
  Floteka. Test potwierdza nową ścieżkę oraz przejęcie wcześniejszej bazy.
- **Rzeczywisty Excel 16.0, plik XLSM** — sprawdzono cztery miejsca,
  szkic, automatyczny zapis pełnego wpisu, deduplikację, zachowanie poprzedniej
  wersji, czyszczenie, przywrócenie, kolejność zajęcia miejsc, zmianę tira,
  ochronę arkusza historii i zapis/ponowne otwarcie.
- **Pięć regresji w rzeczywistym Excelu** — wymuszony błąd tworzenia
  kierowcy cofa częściowy blok; błąd jednego wiersza wklejenia nie blokuje
  poprawnej zmiany drugiego kierowcy; błędna data przywraca poprzednie dane;
  brakująca kopia jest odtwarzana z pełnego wpisu; brak technicznego indeksu
  nie tworzy duplikatu istniejącej historii.
- **Filtry i wydajność** — 250 kierowców, 1000 kompletnych miejsc i 1000
  rekordów historii. Filtrowanie, zapis i ponowne otwarcie pozostawiły dokładnie
  1000 rekordów. Poprawka dodała jedną nową wersję. Każda wcześniejsza komórka
  pozostała identyczna. Sprawdzono także pusty wynik filtra, dodanie kierowcy
  pod filtrem, wielokrotny wybór i zakres dat. Wszystkie klucze końcowego testu
  były unikalne. Mediana trzech zapisów: wcześniej 6,8059 s, po poprawce
  0,3319 s (lokalny Excel 16.0; niezmienione dane; aktualna próba z polami opcjonalnymi).
- **Zachowanie danych użytkownika** — poprawiona kopia zmienionego skoroszytu
  zachowała wszystkie komórki danych, rejestrów, szkiców i historii.
  Suma SHA-256 oryginału pozostała bez zmian.
- **Gotowy EXE** — uruchomiono spakowany program, rzeczywisty formularz Tk,
  zapis SQLite, historię, zwalnianie miejsca i eksport XLSX na izolowanej bazie.
- **Podgląd GUI** — sprawdzono układ na obrazie pobranym bezpośrednio
  z okna aplikacji, z fikcyjnymi danymi. Podgląd nie zawiera pulpitu użytkownika.
- **Przegląd niezależnego recenzenta** — cztery istotne uwagi dostały
  testy reprodukujące problem; po poprawkach testy przeszły. Dodatkowy przegląd
  poprawek Excela wskazał potrzebę kontroli brakującej historii przy szybkim
  zapisie i korekty domyślnej ścieżki testów; oba punkty poprawiono.

## Zakres i ograniczenia

Testy uruchomiono lokalnie na Windows z Pythonem 3.13.15 i Excelem 16.0.
Nie testowano innych wersji systemu, LibreOffice, Excela przeglądarkowego
ani równoczesnej pracy wielu użytkowników. Narzędzia są przeznaczone dla
jednego stanowiska i lokalnego dysku.

Ochrona historii w skoroszycie zapobiega przypadkowej edycji; nie jest
zabezpieczeniem kryptograficznym. Również plik bazy można usunąć narzędziami
systemowymi, dlatego zalecane jest wykonywanie kopii przyciskiem w programie.

W lokalnym środowisku nie było dostępnej drukarki: nie potwierdzono wydruku
na papierze. Układ ma orientację poziomą, nagłówki i dopasowanie do szerokości.
Excel odrzucił programowe ustawienie A4 bez drukarki, więc pozostawiono jego
domyślny format papieru.

Makra należy włączyć dla używanego pliku. Do działania gotowego XLSM nie jest
potrzebny programistyczny dostęp do projektu VBA. Podczas budowy przywrócono
poprzednią wartość AccessVBOM.

Surowe raporty: `docs/excel-test.json`, `docs/excel-regressions.json`,
`docs/exe-test.json`,
`docs/excel-performance-after.json`, `docs/excel-migration.json`
w archiwum źródeł. Skrypt `tools/build_release.ps1`
powtarza testy, budowę i pakowanie.

Raporty nowych funkcji: `docs/excel-deadlines.json`, `docs/excel-optional.json`.
Niezależny przegląd wykrył cztery istotne problemy opisane powyżej;
poprawiono je i zweryfikowano testami regresji.

Eksporty utworzone starszą wersją aplikacji należy wygenerować ponownie.
