import pandas as pd
import glob
import os
import matplotlib.pyplot as plt
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.chart import BarChart, LineChart, Reference

# ========================================
# KONFIGURACJA
# ========================================

# Ścieżka do folderu z danymi
SCIEZKA_DANYCH = r'C:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\Polregio'

# Lista przewoźników
PRZEWOZNICY = ['IC', 'PR', 'KW', 'ARP', 'SKM', 'KD', 'KS', 'LKA', 'Suma']

# Lista typów danych
TYPY_DANYCH = ['Brutto', 'Karta', 'BLIK', 'Netto', 'Prowizja', 'Ilość']

# ========================================
# FUNKCJE POMOCNICZE
# ========================================

def wczytaj_excel(plik):
    """Wczytuje plik Excel z automatycznym doborem silnika"""
    try:
        return pd.read_excel(plik, sheet_name=0, engine='openpyxl')
    except:
        try:
            return pd.read_excel(plik, sheet_name=0, engine='xlrd')
        except Exception as e:
            raise Exception(f"Nie można wczytać pliku: {e}")


def pobierz_automaty(df):
    """Pobiera listę numerów automatów z pierwszej kolumny"""
    if df.empty:
        return []
    return sorted([str(x).strip() for x in df.iloc[:, 0].unique().tolist()])


def pobierz_dostepne_przewozniki(df):
    """Wykrywa dostępnych przewoźników w kolumnach DataFrame"""
    dostepni = []
    for col in df.columns:
        parts = str(col).split()
        if len(parts) >= 2 and parts[1] in PRZEWOZNICY:
            przewoznik = parts[1]
            if przewoznik not in dostepni:
                dostepni.append(przewoznik)
    return dostepni


def pobierz_dostepne_typy(df):
    """Wykrywa dostępne typy danych w kolumnach DataFrame"""
    dostepne = []
    for col in df.columns:
        parts = str(col).split()
        if len(parts) >= 1:
            typ = parts[0]
            if typ in TYPY_DANYCH and typ not in dostepne:
                dostepne.append(typ)
    return dostepne


def tworz_zestawienie(df, automaty, przewoznicy, typy):
    """
    Tworzy zestawienie danych dla wybranych parametrów
    Wiersze: automaty
    Kolumny: typy danych × przewoźnicy (jeśli więcej niż jeden przewoźnik)
    """
    zestawienie_lista = []
    
    # Filtruj po wybranych automatach
    df_filtr = df[df.iloc[:, 0].astype(str).isin(automaty)]
    
    for _, row in df_filtr.iterrows():
        numer_automatu = str(row.iloc[0])
        dane_wiersza = {'Nr automatu': numer_automatu}
        
        # Dla każdego przewoźnika i każdego typu danych
        for przewoznik in przewoznicy:
            for typ in typy:
                # Szukaj kolumny: "Typ Przewoźnik"
                szukana_kolumna = f"{typ} {przewoznik}"
                if szukana_kolumna in df.columns:
                    wartosc = row[szukana_kolumna]
                    try:
                        wartosc = float(wartosc) if pd.notna(wartosc) else 0
                    except:
                        wartosc = 0
                    # Jeśli jest tylko jeden przewoźnik, kolumna to tylko typ
                    # Jeśli jest więcej, kolumna to "Typ (Przewoźnik)"
                    if len(przewoznicy) == 1:
                        nazwa_kolumny = typ
                    else:
                        nazwa_kolumny = f"{typ} ({przewoznik})"
                    dane_wiersza[nazwa_kolumny] = wartosc
                else:
                    if len(przewoznicy) == 1:
                        nazwa_kolumny = typ
                    else:
                        nazwa_kolumny = f"{typ} ({przewoznik})"
                    dane_wiersza[nazwa_kolumny] = 0
        
        zestawienie_lista.append(dane_wiersza)
    
    if not zestawienie_lista:
        return None
    
    df_zestawienie = pd.DataFrame(zestawienie_lista)
    df_zestawienie.set_index('Nr automatu', inplace=True)
    return df_zestawienie


def zapisz_wykres_png(zestawienie_df, sciezka_zapisu):
    """Tworzy i zapisuje wykres jako PNG"""
    num_rows = len(zestawienie_df)
    
    plt.figure(figsize=(12, 7))
    
    if num_rows > 5:
        # Wykres liniowy dla wielu automatów
        for index, row in zestawienie_df.iterrows():
            plt.plot(zestawienie_df.columns, row.values, marker='o', label=str(index))
        plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        plt.title("Wykres liniowy danych")
    else:
        # Wykres słupkowy dla kilku automatów
        zestawienie_df.T.plot(kind='bar', ax=plt.gca(), width=0.8)
        plt.title("Wykres słupkowy danych")
        plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    
    plt.xlabel("Typ danych")
    plt.ylabel("Wartość")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(sciezka_zapisu, dpi=150, bbox_inches='tight')
    plt.close()


def eksportuj_do_excela(zestawienie_df, sciezka_wykresu, sciezka_wyjsciowa):
    """Eksportuje zestawienie i wykres do nowego pliku Excel"""
    # Zapisz zestawienie
    zestawienie_df.to_excel(sciezka_wyjsciowa, sheet_name='Analiza', index=True)
    
    # Dodaj wykres do Excela
    if os.path.exists(sciezka_wykresu):
        wb = load_workbook(sciezka_wyjsciowa)
        ws = wb.active
        
        img = XLImage(sciezka_wykresu)
        img.width = 600
        img.height = 400
        
        # Wstaw wykres obok zestawienia
        kolumna_wykresu = chr(65 + len(zestawienie_df.columns) + 2)  # 2 kolumny dalej
        ws.add_image(img, f'{kolumna_wykresu}2')
        
        wb.save(sciezka_wyjsciowa)
    
    return sciezka_wyjsciowa

# ========================================
# ETAP 1: WCZYTYWANIE PLIKÓW
# ========================================

print("=" * 60)
print("ETAP 1: WCZYTYWANIE PLIKÓW EXCEL")
print("=" * 60)

# Znajdź wszystkie pliki Excel w folderze
pliki_excel = glob.glob(os.path.join(SCIEZKA_DANYCH, "*.xls*"))

if not pliki_excel:
    print(f"❌ Nie znaleziono plików Excel w folderze: {SCIEZKA_DANYCH}")
    exit()

print(f"✅ Znaleziono {len(pliki_excel)} plików Excel:")
for plik in pliki_excel:
    print(f"   - {os.path.basename(plik)}")

# Wczytaj pierwszy plik (można rozszerzyć na więcej plików)
try:
    df_dane = wczytaj_excel(pliki_excel[0])
    print(f"\n✅ Wczytano plik: {os.path.basename(pliki_excel[0])}")
    print(f"   Wierszy: {len(df_dane)}, Kolumn: {len(df_dane.columns)}")
except Exception as e:
    print(f"❌ Błąd przy wczytywaniu: {e}")
    exit()


# ========================================
# ETAP 2: WYBÓR DANYCH
# ========================================

print("\n" + "=" * 60)
print("ETAP 2: WYBÓR DANYCH")
print("=" * 60)

# 2a) Analiza według PRZEWOŹNIKA
print("\n2a) WYBÓR PRZEWOŹNIKÓW")
print("-" * 40)

dostepni_przewoznicy = pobierz_dostepne_przewozniki(df_dane)
if not dostepni_przewoznicy:
    print("❌ Nie znaleziono przewoźników w danych")
    exit()

print("Dostępni przewoźnicy:")
for i, przewoznik in enumerate(dostepni_przewoznicy, 1):
    print(f"   {i}. {przewoznik}")

print("\nWybierz przewoźników (np. '1,3,5' dla kilku lub 'wszystkie'):")
while True:
    wybor_przewoznik = input("Wybór: ").strip().lower()
    
    if wybor_przewoznik == 'wszystkie':
        wybrani_przewoznicy = dostepni_przewoznicy.copy()
        print(f"✅ Wybrano wszystkich przewoźników: {', '.join(wybrani_przewoznicy)}")
        break
    else:
        try:
            indeksy = [int(x.strip()) - 1 for x in wybor_przewoznik.split(',')]
            wybrani_przewoznicy = [dostepni_przewoznicy[i] for i in indeksy if 0 <= i < len(dostepni_przewoznicy)]
            if wybrani_przewoznicy:
                print(f"✅ Wybrano {len(wybrani_przewoznicy)} przewoźników: {', '.join(wybrani_przewoznicy)}")
                break
            else:
                print("⚠️ Nieprawidłowe numery, spróbuj ponownie")
        except (ValueError, IndexError):
            print("⚠️ Nieprawidłowy format, użyj np. '1,2,3' lub 'wszystkie'")

# 2b) Wybór numeru automatu
print("\n2b) WYBÓR NUMERÓW AUTOMATÓW")
print("-" * 40)

dostepne_automaty = pobierz_automaty(df_dane)
if not dostepne_automaty:
    print("❌ Nie znaleziono automatów w danych")
    exit()

print(f"Dostępne automaty ({len(dostepne_automaty)}):")
for automat in dostepne_automaty:
    print(f"   - {automat}")

while True:
    wybor_automaty = input("\nWpisz numery automatów (oddzielone przecinkiem): ").strip()
    wybrane_automaty = [a.strip() for a in wybor_automaty.split(',')]
    
    # Sprawdź czy wszystkie automaty istnieją
    nieprawidlowe = [a for a in wybrane_automaty if a not in dostepne_automaty]
    if nieprawidlowe:
        print(f"⚠️ Nie znaleziono automatów: {', '.join(nieprawidlowe)}")
        print("   Spróbuj ponownie")
    else:
        print(f"✅ Wybrano {len(wybrane_automaty)} automatów: {', '.join(wybrane_automaty)}")
        break

# 2c) Typ danych - wszystkie po kolei
print("\n2c) WYBÓR TYPÓW DANYCH")
print("-" * 40)

dostepne_typy = pobierz_dostepne_typy(df_dane)
if not dostepne_typy:
    print("❌ Nie znaleziono typów danych")
    exit()

print("Dostępne typy danych:")
for i, typ in enumerate(dostepne_typy, 1):
    print(f"   {i}. {typ}")

print("\nWybierz typy danych (np. '1,3,5' lub 'wszystkie'):")
while True:
    wybor_typy = input("Wybór: ").strip().lower()
    
    if wybor_typy == 'wszystkie':
        wybrane_typy = dostepne_typy.copy()
        print(f"✅ Wybrano wszystkie typy: {', '.join(wybrane_typy)}")
        break
    else:
        try:
            indeksy = [int(x.strip()) - 1 for x in wybor_typy.split(',')]
            wybrane_typy = [dostepne_typy[i] for i in indeksy if 0 <= i < len(dostepne_typy)]
            if wybrane_typy:
                print(f"✅ Wybrano typy: {', '.join(wybrane_typy)}")
                break
            else:
                print("⚠️ Nieprawidłowe numery, spróbuj ponownie")
        except (ValueError, IndexError):
            print("⚠️ Nieprawidłowy format, użyj np. '1,2,3' lub 'wszystkie'")


# ========================================
# ETAP 3: TWORZENIE ZESTAWIENIA I WYKRESU
# ========================================

print("\n" + "=" * 60)
print("ETAP 3: TWORZENIE ZESTAWIENIA I WYKRESU")
print("=" * 60)

# Tworzenie zestawienia
print("\n📊 Tworzenie zestawienia danych...")
zestawienie = tworz_zestawienie(df_dane, wybrane_automaty, wybrani_przewoznicy, wybrane_typy)

if zestawienie is None or zestawienie.empty:
    print("❌ Nie udało się utworzyć zestawienia")
    exit()

print("✅ Zestawienie utworzone:")
print(zestawienie)

# Generowanie nazwy pliku z timestamp
timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
if len(wybrani_przewoznicy) == 1:
    nazwa_pliku = f"Analiza_{wybrani_przewoznicy[0]}_{timestamp}"
else:
    nazwa_pliku = f"Analiza_{len(wybrani_przewoznicy)}przewoznikow_{timestamp}"

# Tworzenie wykresu PNG
sciezka_wykresu = os.path.join(SCIEZKA_DANYCH, f"{nazwa_pliku}.png")
print(f"\n📈 Tworzenie wykresu...")
zapisz_wykres_png(zestawienie, sciezka_wykresu)
print(f"✅ Wykres zapisany: {nazwa_pliku}.png")

# Eksport do Excela
sciezka_excel = os.path.join(SCIEZKA_DANYCH, f"{nazwa_pliku}.xlsx")
print(f"\n💾 Eksportowanie do Excela...")
try:
    eksportuj_do_excela(zestawienie, sciezka_wykresu, sciezka_excel)
    print(f"✅ Plik Excel zapisany: {nazwa_pliku}.xlsx")
    
    # Otwórz plik (Windows)
    try:
        os.startfile(sciezka_excel)
        print("✅ Plik otwarty w Excelu")
    except:
        print("⚠️ Nie udało się automatycznie otworzyć pliku")
        
except Exception as e:
    print(f"❌ Błąd podczas eksportu: {e}")

print("\n" + "=" * 60)
print("✅ PROGRAM ZAKOŃCZONY POMYŚLNIE")
print("=" * 60)