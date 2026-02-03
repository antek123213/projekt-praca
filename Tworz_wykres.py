#wcytywanie plikow do pythona
import pandas as pd
import glob
import os
import matplotlib.pyplot as plt

# 1. Ścieżka do folderu
sciezka = r'C:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\PKP IC\2024\test'

# 2. Pobranie wszystkich plików .xls (możesz użyć "*.xls*" aby złapać też .xlsx)
pliki_xls = glob.glob(os.path.join(sciezka, "*.xls*"))

# 3. Słownik na dane
zmienne_excel = {}

# Lista przewoźników (hierarchia/priorytet)
PRZEWOZNICY = ['IC', 'PR', 'KW', 'ARP', 'SKM', 'KD', 'KS', 'KS', 'LKA']

# Funkcja pomocnicza do podgrupowania kolumn po przewoźniku (2. wyraz)
def podgrupuj_kolumny_po_przewozniku(df, przewoznicy=None):
    """Tworzy układ długi: drugi wyraz w nazwie kolumny traktuje jako przewoźnika"""
    if przewoznicy is None:
        przewoznicy = []
    cols = list(df.columns)
    split_cols = []
    id_cols = []
    parsed = {}

    for col in cols:
        parts = str(col).split()
        if len(parts) >= 2:
            przewoznik = parts[1]
            if not przewoznicy or przewoznik in przewoznicy:
                miara = parts[0]
                split_cols.append(col)
                parsed[col] = (miara, przewoznik)
        else:
            id_cols.append(col)

    if not split_cols:
        return df

    rows = []
    for col in split_cols:
        miara, przewoznik = parsed[col]
        if id_cols:
            tmp = df[id_cols].copy()
        else:
            tmp = pd.DataFrame(index=df.index)
        tmp['Miara'] = miara
        tmp['Przewoźnik'] = przewoznik
        tmp['Wartość'] = df[col]
        rows.append(tmp)

    return pd.concat(rows, ignore_index=True)

# Funkcja do wyboru kolumn zawierających "Suma"
def wybierz_kolumny_po_sumie(df):
    """Wybiera kolumny zawierające słowo 'Suma'"""
    cols = list(df.columns)
    wybrane = []
    
    for col in cols:
        if "Suma" in str(col):
            wybrane.append(col)
    
    if not wybrane:
        return None
    
    typy = [str(c).split()[0] for c in wybrane]  # Pierwsze słowo
    print(f"✅ Znaleziono {len(wybrane)} kolumn Suma: {', '.join(typy)}")
    return wybrane

# Funkcja do wyboru kolumn na podstawie listy przewoźników
def wybierz_kolumny_po_przewozniku(df, przewoznicy):
    """Wybiera kolumny, których 2. wyraz jest na liście przewoźników"""
    cols = list(df.columns)
    wybrane = []
    
    for col in cols:
        parts = str(col).split()
        if len(parts) >= 2 and parts[1] in przewoznicy:
            wybrane.append(col)
    
    if not wybrane:
        return None
    
    print(f"✅ Znaleziono {len(wybrane)} kolumn dla przewoźników: {', '.join(set(str(c).split()[1] for c in wybrane if len(str(c).split()) >= 2))}")
    return wybrane

# Funkcja do automatycznego wyboru enginu
def wczytaj_excel(plik, usecols=None):
    """Próbuje wczytać plik Excel, automatycznie dobiera silnik"""
    try:
        # Spróbuj z openpyxl (dla .xlsx)
        return pd.read_excel(plik, sheet_name=0, usecols=usecols, engine='openpyxl')
    except Exception as e1:
        try:
            # Spróbuj z xlrd (dla .xls)
            return pd.read_excel(plik, sheet_name=0, usecols=usecols, engine='xlrd')
        except Exception as e2:
            raise Exception(f"Nie mogę wczytać pliku żadnym silnikiem: openpyxl={e1}, xlrd={e2}")

# ETAP 1: WCZYTYWANIE PLIKÓW EXCEL
if pliki_xls:
    for plik in pliki_xls:
        nazwa_zmiennej = os.path.splitext(os.path.basename(plik))[0]
        try:
            df_temp = wczytaj_excel(plik)
            # Sprawdź czy są kolumny ze słowem "Suma"
            wybrane_kolumny = wybierz_kolumny_po_sumie(df_temp)
            if wybrane_kolumny is None:
                # Jeśli nie ma "Suma", szukaj po przewoźnikach
                wybrane_kolumny = wybierz_kolumny_po_przewozniku(df_temp, PRZEWOZNICY)
            if wybrane_kolumny is None:
                # Jeśli nadal nic nie znaleziono, wczytaj wszystko
                wybrane_kolumny = list(df_temp.columns)
            # ZAWSZE dołącz pierwszą kolumnę (Nr aut.)
            pierwsza_kolumna = df_temp.columns[0]
            if pierwsza_kolumna not in wybrane_kolumny:
                wybrane_kolumny = [pierwsza_kolumna] + wybrane_kolumny
            df = wczytaj_excel(plik, usecols=wybrane_kolumny)
            zmienne_excel[nazwa_zmiennej] = df
        except Exception as e:
            print(f"❌ Błąd przy pliku {nazwa_zmiennej}: {e}")
else:
    print(f"❌ Nie znaleziono plików Excel w folderze: {sciezka}")
    exit()


# ETAP 2: PODGRUPOWANIE DANYCH PO PRZEWOŹNIKACH

# Tworzenie listy nazw wczytanych plików
lista_plików = list(zmienne_excel.keys())

# Pobranie dostępnych automatów i wykrycie dostępnych struktur danych
df_temp = zmienne_excel.get(lista_plików[0], pd.DataFrame()) if lista_plików else pd.DataFrame()
# Pobierz pierwszą kolumnę jako listę automatów (kolumna A w Excelu)
if not df_temp.empty:
    # Konwertuj do stringów aby uniknąć problemów z typami (int vs str)
    dostepne_automaty = sorted([str(x).strip() for x in df_temp.iloc[:, 0].unique().tolist()])
else:
    dostepne_automaty = []

# Sprawdź jakie struktury są dostępne
ma_kolumny_suma = False
dostepne_typy = []
for col in df_temp.columns:
    if "Suma" in str(col):
        ma_kolumny_suma = True
        typ = str(col).split()[0]
        if typ not in dostepne_typy:
            dostepne_typy.append(typ)

ma_kolumny_przewoznikow = False
for col in df_temp.columns:
    parts = str(col).split()
    if len(parts) >= 2 and parts[1] in PRZEWOZNICY:
        ma_kolumny_przewoznikow = True
        break

print("\n📋 DOSTĘPNE AUTOMATY:", dostepne_automaty)

# Pytanie o tryb pracy - zawsze pokazuj obie opcje
print("\n🔍 WYBÓR TRYBU ANALIZY:")
print("1. Analiza według TYPU (Brutto, Karta, BLIK, Netto, Prowizja, Ilość)")
print("2. Analiza według PRZEWOŹNIKA (IC, PR, KW, ARP, SKM, KD, KS, LKA)")

wybor_trybu = input("\nWybierz tryb (1 lub 2): ").strip()

while wybor_trybu not in ["1", "2"]:
    print(f"⚠️ Nieprawidłowy wybór. Wpisz 1 lub 2")
    wybor_trybu = input("Wybierz tryb (1 lub 2): ").strip()

if wybor_trybu == "1":
    tryb = "typy"
    if not ma_kolumny_suma:
        print("⚠️ Uwaga: Nie znaleziono kolumn ze słowem 'Suma'. Program spróbuje dopasować dane.")
    print("\n📋 DOSTĘPNE TYPY DANYCH:", dostepne_typy if dostepne_typy else "Brak")
else:
    tryb = "przewoznicy"
    if ma_kolumny_suma and not ma_kolumny_przewoznikow:
        print("\n⚠️⚠️⚠️ UWAGA! ⚠️⚠️⚠️")
        print("Wykryto kolumny ze słowem 'Suma' (Brutto Suma, Karta Suma, itd.)")
        print("Dla tego typu danych powinieneś wybrać opcję 1 (Analiza według TYPU)!")
        print("\nChcesz kontynuować mimo to? (tak/nie)")
        kontynuuj = input().strip().lower()
        if kontynuuj not in ['tak', 't', 'yes', 'y']:
            print("Program zakończony. Uruchom ponownie i wybierz opcję 1.")
            exit()
    print("\n📋 DOSTĘPNI PRZEWOŹNICY:", PRZEWOZNICY)

# Pytanie o automat
print("\nWybór automatu:")
print("Wpisz numer automatu (lub kilka oddzielone przecinkiem) z listy powyżej")
automaty_input = input("Numer automatu: ").strip()

# Parsowanie automatów (może być jeden lub kilka oddzielonych przecinkiem)
wybrane_automaty = [str(a).strip() for a in automaty_input.split(',')]

# Walidacja automatów
nieprawidlowe = [a for a in wybrane_automaty if a not in dostepne_automaty]

if nieprawidlowe:
    print(f"⚠️ Automaty nie znalezione: {nieprawidlowe}")
    print(f"Dostępne: {dostepne_automaty}")
    automaty_input = input("Podaj prawidłowe numery automatów (oddzielone przecinkiem): ").strip()
    wybrane_automaty = [str(a).strip() for a in automaty_input.split(',')]

print(f"✅ Wybrano automatów: {', '.join(wybrane_automaty)}")

# Pytanie zależne od struktury danych
if tryb == "typy":
    # Struktura ze słowem "Suma" - pytaj o typy
    print("\nWybór typu danych:")
    print("Wpisz typ z listy powyżej (np. Brutto, Karta, BLIK, Netto, Prowizja, Ilość)")
    print("Możesz podać kilka typów oddzielonych przecinkiem")
    jaki_wybor = input("Typ: ").strip()
    
    # Parsowanie i walidacja typów
    wybrane_opcje = [p.strip() for p in jaki_wybor.split(',')]
    
    # Sprawdź czy wszystkie typy są na liście
    nieprawidlowe = [p for p in wybrane_opcje if p not in dostepne_typy]
    
    if nieprawidlowe:
        print(f"⚠️ Typy nie znalezione: {nieprawidlowe}")
        print(f"Dostępne: {dostepne_typy}")
        jaki_wybor = input("Podaj prawidłowe nazwy typów (oddzielone przecinkiem): ").strip()
        wybrane_opcje = [p.strip() for p in jaki_wybor.split(',')]
    
    print(f"✅ Wybrano typów: {', '.join(wybrane_opcje)}")
else:
    # Struktura z przewoźnikami - pytaj o przewoźników
    print("\nWybór przewoźnika:")
    print("Wpisz nazwę przewoźnika z listy powyżej ")
    jaki_wybor = input("Przewoźnik: ").strip()
    
    # Parsowanie i walidacja przewoźników
    if jaki_wybor.lower() != 'ogólny':
        wybrane_opcje = [p.strip() for p in jaki_wybor.split(',')]
        
        # Sprawdź czy wszystkie przewoźnicy są na liście
        nieprawidlowi = [p for p in wybrane_opcje if p not in PRZEWOZNICY]
        
        if nieprawidlowi:
            print(f"⚠️ Przewoźnicy nie znalezieni: {nieprawidlowi}")
            print(f"Dostępni: {PRZEWOZNICY}")
            jaki_wybor = input("Podaj prawidłowe nazwy przewoźników (oddzielone przecinkiem) lub 'ogólny': ").strip()
            if jaki_wybor.lower() != 'ogólny':
                wybrane_opcje = [p.strip() for p in jaki_wybor.split(',')]
        
        print(f"✅ Wybrano przewoźników: {', '.join(wybrane_opcje)}")
    else:
        wybrane_opcje = PRZEWOZNICY  # Wszyscy przewoźnicy
        print(f"✅ Wybrano wszystkich przewoźników")

# Inicjalizuj zmienną zestawienia
zestawienie = None

def tworz_zestawienie_excel(df, wybrane_opcje, tryb):
    """Tworzy zestawienie: wiersze=automaty, kolumny=typy lub przewoźnicy (zależnie od trybu)"""
    
    if df.empty:
        return None
    
    # Budowanie zestawienia dla WSZYSTKICH automatów w DataFramie
    zestawienie_lista = []
    
    # Dla każdego automatu (wiersza)
    for idx, row in df.iterrows():
        numer_automatu = str(row.iloc[0])  # Pierwsza kolumna to numer automatu
        zestawienie_dane = {}
        
        # Przeglądaj wszystkie kolumny
        for col in df.columns:
            parts = str(col).split()
            
            if tryb == "typy" and "Suma" in str(col) and len(parts) >= 1:
                # Tryb typów - szukaj kolumn ze "Suma", bierz pierwsze słowo
                typ = parts[0]
                if typ in wybrane_opcje:
                    wartosc = row[col]
                    try:
                        wartosc = float(wartosc) if pd.notna(wartosc) else 0
                    except:
                        wartosc = 0
                    zestawienie_dane[typ] = wartosc
                    
            elif tryb == "przewoznicy" and len(parts) >= 2:
                # Tryb przewoźników - szukaj po drugim słowie
                przewoznik = parts[1]
                if przewoznik in wybrane_opcje:
                    wartosc = row[col]
                    try:
                        wartosc = float(wartosc) if pd.notna(wartosc) else 0
                    except:
                        wartosc = 0
                    zestawienie_dane[przewoznik] = wartosc
        
        if zestawienie_dane:
            zestawienie_lista.append((numer_automatu, zestawienie_dane))
    
    if not zestawienie_lista:
        print("⚠️ Brak danych do zestawienia")
        return None
    
    # Utwórz DataFrame z zestawieniem
    zestawienie_df = pd.DataFrame([dane for _, dane in zestawienie_lista], 
                                   index=[aut for aut, _ in zestawienie_lista])
    zestawienie_df.index.name = 'Nr aut.'
    
    return zestawienie_df

# Funkcja do przetwarzania danych i rysowania wykresu
def przetwarzaj_dane_i_rysuj(lista_plikow, lista_automatow, wybrane_opcje, tryb, sciezka_wykresu):
    """Pobiera dane, tworzy zestawienie i rysuje wykres dla listy automatów"""
    global zestawienie
    
    df_dane = zmienne_excel.get(lista_plikow[0], None)
    if df_dane is None:
        print("❌ Brak danych do wykresu")
        return None
    
    # Konwertuj do stringów dla pewności
    lista_automatow = [str(a) for a in lista_automatow]
    
    # Filtrowanie dla wybranych automatów
    df_dane = df_dane[df_dane.iloc[:, 0].astype(str).isin(lista_automatow)]
    print(f"✅ Pobrano dane dla automatów: {', '.join(lista_automatow)}")
    
    if df_dane.empty:
        print("⚠️ Brak danych do zestawienia")
        return None
    
    # Tworzenie zestawienia
    zestawienie = tworz_zestawienie_excel(df_dane, wybrane_opcje, tryb)
    if zestawienie is None:
        print("⚠️ Brak danych do zestawienia")
        return None
    
    print("\n📊 Zestawienie:")
    print(zestawienie)
    
    # Rysowanie wykresu i zapis do pliku
    rysuj_wykres(zestawienie, sciezka_wykresu)
    print(f"✅ Wykres zapisany: {sciezka_wykresu}")
    
    return zestawienie

# Funkcja do rysowania wykresu
def rysuj_wykres(zestawienie_df, sciezka_zapisania):
    """Rysuje wykres na podstawie liczby wierszy i zapisuje do PNG"""
    
    # Liczba wierszy (automatów)
    num_rows = len(zestawienie_df)
    
    # Wybierz typ wykresu na podstawie liczby wierszy
    if num_rows > 5:
        typ_wykresu = "liniowy"
    else:
        typ_wykresu = "słupkowy"
    
    plt.figure(figsize=(10, 6))
    
    if typ_wykresu == "liniowy":
        # Wykres liniowy
        for index, row in zestawienie_df.iterrows():
            plt.plot(zestawienie_df.columns, row.values, marker='o', label=str(index))
        plt.legend()
        plt.title(f"Wykres liniowy")
    else:
        # Wykres słupkowy
        zestawienie_df.T.plot(kind='bar', ax=plt.gca())
        plt.title(f"Wykres słupkowy")
    
    plt.xlabel("Typy danych")
    plt.ylabel("Wartość")
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    # Zapisz wykres
    plt.savefig(sciezka_zapisania, dpi=100, bbox_inches='tight')
    plt.close()

# Wywołanie funkcji przetwarzania
timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
sciezka_wykresu = os.path.join(sciezka, f"Wykres_{timestamp}.png")
zestawienie = przetwarzaj_dane_i_rysuj(lista_plików, wybrane_automaty, wybrane_opcje, tryb, sciezka_wykresu)

# ========================================
# ETAP 3: EXPORT DO PLIKU EXCEL
# ========================================

# Jeśli zestawienie zostało utworzone, zapisz do pliku Excel i otwórz
if zestawienie is not None:
    # Wygeneruj nazwę pliku
    nazwa_pliku = f"Zestawienie_{timestamp}.xlsx"
    sciezka_wyjsciowa = os.path.join(sciezka, nazwa_pliku)
    
    # Zapisz do Excela
    try:
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image as XLImage
        
        # Najpierw zapisz zestawienie
        zestawienie.to_excel(sciezka_wyjsciowa, sheet_name='Zestawienie', index=True)
        print(f"\n✅ Zestawienie zapisane do: {sciezka_wyjsciowa}")
        
        # Teraz dodaj wykres do Excela
        if os.path.exists(sciezka_wykresu):
            wb = load_workbook(sciezka_wyjsciowa)
            ws = wb.active
            
            # Wstaw obraz wykresu obok zestawienia (kolumna do prawej)
            img = XLImage(sciezka_wykresu)
            img.width = 400
            img.height = 300
            
            # Wstaw w kolumnie F (obok zestawienia)
            ws.add_image(img, 'F2')
            
            wb.save(sciezka_wyjsciowa)
            print(f"✅ Wykres wstawiony do Excela")
        
        # Otwórz plik w Excelu (tylko Windows)
        try:
            os.startfile(sciezka_wyjsciowa)
            print("✅ Plik otwarty w Excelu")
        except Exception as e:
            print(f"⚠️ Nie udało się otworzyć pliku automatycznie: {e}")
    except Exception as e:
        print(f"❌ Błąd przy zapisywaniu do Excela: {e}")
else:
    print("\n⚠️ Brak zestawienia do eksportu")