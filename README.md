# 📊 Excel Makro: Połącz tekst w jednej komórce

To proste makro VBA umożliwia szybkie łączenie tekstu z wielu zaznaczonych komórek w Excelu. Wynik pojawia się w pierwszej (lewa-górna) komórce zaznaczenia, a pozostałe komórki zostają wyczyszczone.

## ✨ Funkcje

- 🔗 Łączy tekst z wielu komórek w jeden ciąg.
- 📍 Wstawia wynik do pierwszej komórki zaznaczenia.
- 🧹 Czyści pozostałe komórki.
- 🕹️ Działa w każdej wersji Excela z obsługą makr (Windows/Mac).

---

## 📦 Zawartość

- `PolaczTekst.bas` – plik modułu VBA z gotowym makrem do importu.

---

## 🛠️ Jak zainstalować

1. Otwórz plik Excela.
2. Naciśnij `Alt + F11`, aby wejść do edytora VBA.
3. Kliknij: `Plik > Importuj plik...` i wybierz `PolaczTekst.bas`.
4. Zamknij edytor.
5. Zaznacz kilka komórek z tekstem.
6. Uruchom makro `PolaczTekstWJednejKomorce` (`Alt + F8` > wybierz z listy).

---

## 📌 Przykład użycia

Zaznaczenie:

| A          | B          | C           |
|------------|------------|-------------|
| Hello      | world!     | :)          |

Po uruchomieniu makra:

| A                   | B   | C   |
|---------------------|-----|-----|
| Hello world! :)     |     |     |

---

## ✅ Wymagania

- Excel z obsługą VBA (np. Excel 2016, 2019, 365, itd.)
- Włączone makra (upewnij się, że plik ma rozszerzenie `.xlsm` lub `.xlsb`)

---

## 📄 Licencja

Projekt open-source. Możesz używać, modyfikować i udostępniać dalej bez ograniczeń 🎉

---

## 💬 Masz pomysł?

Chcesz dodać separator (np. przecinek, nową linię)? Masz inne potrzeby? Stwórz issue lub forka – chętnie pomogę rozbudować makro! 😄
