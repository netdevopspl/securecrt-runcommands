# **RUN AND COMPARE COMMANDS**

- [**RUN AND COMPARE COMMANDS**](#run-and-compare-commands)
  - [**WSTĘP**](#wstęp)
  - [**OPIS TECHNICZNY**](#opis-techniczny)
    - [LISTA POLECEŃ](#lista-poleceń)
    - [JAK DZIAŁA SKRYPT](#jak-działa-skrypt)
    - [KOMUNIKATY](#komunikaty)
  - [**WYMAGANIA**](#wymagania)
  - [**JAK ZACZĄĆ**](#jak-zacząć)
  - [**PROBLEMY**](#problemy)
  - [**KONTAKT**](#kontakt)
  - [**LICENCJA**](#licencja)

---

## **WSTĘP**

Niniejszy skrypt służy do wykonania serii polecen diagnostycznych typu `show`, przed i po zakończeniu dokonywania zmian na urządzeniach Cisco. Pliki te są na koniec porównywane, w celu potwierdzenia poprawności działania sieci.

## **OPIS TECHNICZNY**

### LISTA POLECEŃ

Skrypt korzysta z plików tekstowych, w których zawarte są polecenia diagnostyczne. W zależności od wybranego rodzaju urządzenia, wykorzystywany jest odpowiedni plik.

**Uwaga!** Nazw plików nie można zmieniać (wymagana jest wtedy zmiana w kodzie, linia 238, zmienna `strCommandsFile`)

Lista poleceń powinna uwzględniać też zachowanie się protokołów, a nie tylko samej konfiguracji. Planowana zmiana w konfiguracji nie zawsze oznacza późniejsze, prawidłowe działanie sieci.

### JAK DZIAŁA SKRYPT

- Przy pierwszym uruchomieniu
  - Wybieramy rodzaj urządzenia
  - Wyłączane jest tymczasowow logowanie (jeżeli było włączone) dla aktualnej sesji, z zachowaniem oryginalnej nazwy pliku
  - Torzony jest nowy plik o nazwie `<hostname>_<data>_before.txt`
  - Na aktywnym urządzeniu wykonywane są polecenia, a ich wynik zapisywany jest w utworzonym wcześniej pliku
  - Odtworzenie oryginalnej ścieżki dla pliku logowania
- Przy drugim uruchomieniu
  - Skrypt traktuje uruchomienie jako drugie jeżeli plik *przed* już istnieje
  - Wybieramy rodzaj urządzenia (**Uwaga!** Typ urządzenia powinien być ten sam co poprzednio)
  - Wyłączane jest tymczasowow logowanie (jeżeli było włączone) dla aktualnej sesji, z zachowaniem oryginalnej nazwy pliku
  - Torzony jest nowy plik o nazwie `<hostname>_<data>_after.txt`
  - Na aktywnym urządzeniu wykonywane są polecenia, a ich wynik zapisywany jest w utworzonym wcześniej pliku
  - Odtworzenie oryginalnej ścieżki dla pliku logowania
  - Uruchamiane jest narzędzie ExamDiff, w celu wyświetlenia obu plików *przed* i *po*

Skrypt odgaduje nazwę hostname z urządzenia, ale nie w każdym przypadku to się udaje (np. klastry ASA), więc prosi on o potwierdzenie nazwy hostname przed uruchomieniem poleceń.

### KOMUNIKATY

- Jeżeli nie jesteśmy zalogowani na urządzenie w aktywnym oknie terminala, to skrypt wyświetli stosowny komunikat i zakończy pracę.
- Jeżeli zdefiniowane w konfiguracji katalogi nie istnieją, to skrypt wyświetli stosowny komunikat i zakończy pracę.

## **WYMAGANIA**

- Skrypt przeznaczony jest do użycia w terminalu [Secure CRT](https://www.vandyke.com/products/securecrt/)
- Skrypt korzysta z narzędzia [ExamDiff](http://www.prestosoft.com/edp_examdiff.asp#download) do porównywania plików *przed* i *po*

## **JAK ZACZĄĆ**

1. Przejrzyj pliki txt z listą poleceń i uzupełnij je wg. własnych wymagań

   **Uwaga!** Nigdy nie wykonuj automatycznie poleceń bez wcześniejszej ich weryfikacji

2. Otworz skrypt RunAndCompareCommands.vbs i ustaw odpowiednie zmienne w sekcji *Editable variables*

   a. **strLogPath**: Ścieżka do katalogu, w którym będą zapisywane logi *przed* i *po*

   `Const strLogPath = "c:\Users\test\Documents\Console_Logs\"`

   b. **strCommandsPath**: Ścieżka do katalogu, w którym znajdują się pliki txt z listą poleceń

   `Const strCommandsPath = "c:\Users\test\Documents\SecureCRT-Scripts\"`

   c. **strDiffFile**: Pełna ścieżka do narzędzia ExamDiff

   `Const strDiffFile = "C:\Program Files (x86)\ExamDiff\ExamDiff.exe"`

   d. **blnDebug**: Włączenie (True) lub wyłączenie (False) dodatkowych komunikatów diagnostycznych

   `Const blnDebug = False`

## **PROBLEMY**

- Po wyświetleniu okna IE z wyborem urządzeń, nie staje się ono aktywne. Konieczne jest ręczne przestawienie na przód ekranu.

## **KONTAKT**

E-mail: [krzysztof@nowoczesnysieciowiec.pl](mailto:krzysztof@nowoczesnysieciowiec.pl?Subject=Projekt%20RunAndCompareCommands)

## **LICENCJA**

Autor dołożył wszelkich starań, aby zawarte tu informacje były rzetelne, ale nie gwarantuje ich poprawności. Autor nie bierze odpowiedzialności za żadne szkody wynikające z wykorzystania zawartych tu informacji i skryptów.

Pliki zawarte w tym projekcie mogą być swobodnie wykorzystywane. Mogą one być też dowolnie modyfikowane, z zachowaniem informacji o źródle.

Kopiowanie i dystrybucja możliwa jest tylko z zachowaniem informacji o źródle.

Zawarte w tym projekcie nazwy produktów i znaki towarowe należą do ich prawowitych właścicieli

(C) [Nowoczesny Sieciowiec](https://nowoczesnysieciowiec.pl "Blog Nowoczesny Sieciowiec"), 2021
