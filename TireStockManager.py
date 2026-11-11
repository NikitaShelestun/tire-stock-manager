import sys
import os
import time
import threading

# FAST START - show animation
class FastProgressBar:
    def __init__(self):
        self.stop = False
        self.current = 0
        self.message = "Starting"
    
    def animate(self):
        frames = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
        idx = 0
        while not self.stop:
            frame = frames[idx % len(frames)]
            bar_length = 30
            filled = int(bar_length * self.current // 100)
            bar = '█' * filled + '░' * (bar_length - filled)
            sys.stdout.write(f'\r[{frame}] [{bar}] {self.current:3d}% - {self.message}')
            sys.stdout.flush()
            idx += 1
            time.sleep(0.05)
    
    def start(self):
        self.stop = False
        self.thread = threading.Thread(target=self.animate)
        self.thread.daemon = True
        self.thread.start()
    
    def update(self, percent, message):
        self.current = percent
        self.message = message
    
    def finish(self, message="Complete"):
        self.stop = True
        if self.thread:
            self.thread.join(timeout=0.5)
        bar = '█' * 30
        sys.stdout.write(f'\r[✓] [{bar}] 100% - {message}\n')
        sys.stdout.flush()

# Start animation
progress = FastProgressBar()
progress.start()
progress.update(1, "Initializing...")
time.sleep(0.1)

# import heavy modules
progress.update(3, "Loading pandas...")
import pandas as pd

progress.update(5, "Loading openpyxl...")
import openpyxl
from openpyxl import load_workbook

progress.update(8, "Loading pickle...")
import pickle

progress.update(10, "Loading other modules...")
import gc

# Get paths
def get_excel_path():
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(application_path, 'backup_tires.xlsx')

def get_cache_path():
    return get_excel_path().replace('.xlsx', '_cache.pkl')

def get_hash_path():
    return get_excel_path().replace('.xlsx', '_hash.txt')

def get_file_hash(filepath):
    try:
        with open(filepath, 'rb') as f:
            return hash(f.read())
    except:
        return None

def sort_tires(df):
    if len(df) == 0:
        return df
    df_sorted = df.copy()
    sector_order = {sector: idx for idx, sector in enumerate(sorted(df_sorted['SECTOR'].unique()))}
    df_sorted['SECTOR_ORDER'] = df_sorted['SECTOR'].map(sector_order)
    df_sorted = df_sorted.sort_values(by=['SECTOR_ORDER', 'FLOOR', 'QUANTITY', 'BRAND'],
                                       ascending=[True, True, False, True])
    df_sorted = df_sorted.drop(columns=['SECTOR_ORDER'])
    df_sorted = df_sorted.reset_index(drop=True)
    df_sorted['ID'] = range(1, len(df_sorted) + 1)
    return df_sorted

def load_excel_with_progress(excel_path, progress):
    """Load Excel file and show real progress"""
    
    progress.update(15, "Opening Excel file")
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active
    
    total_rows = ws.max_row - 1
    if total_rows <= 0:
        total_rows = 1
    
    progress.update(20, f"Found {total_rows} rows")
    time.sleep(0.2)
    
    # Read headers
    progress.update(22, "Reading headers")
    headers = []
    for cell in ws[1]:
        headers.append(cell.value)
    
    # Read data row by row
    data = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        percent = 22 + int((row_idx / total_rows) * 70)
        if row_idx % 10 == 0 or row_idx == total_rows:
            progress.update(percent, f"Loading row {row_idx} of {total_rows}")
        
        row_data = {}
        for idx, header in enumerate(headers):
            if idx < len(row):
                row_data[header] = row[idx]
            else:
                row_data[header] = None
        data.append(row_data)
    
    # Convert to DataFrame
    progress.update(93, "Converting to DataFrame")
    df = pd.DataFrame(data)
    
    # Clean data
    progress.update(95, "Cleaning data")
    df = df.fillna('')
    
    numeric_cols = ['ID', 'WIDTH', 'PROFILE', 'DIAMETER', 'FLOOR', 'QUANTITY']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    
    # Sort
    progress.update(98, "Sorting data")
    df = sort_tires(df)
    
    wb.close()
    
    progress.update(100, f"Loaded {len(df)} tires")
    return df

def create_empty_excel():
    excel_path = get_excel_path()
    if not os.path.exists(excel_path):
        empty_df = pd.DataFrame(columns=['ID', 'BRAND', 'WIDTH', 'PROFILE', 'DIAMETER', 
                                          'SEASON', 'CONDITION', 'SECTOR', 'FLOOR', 'QUANTITY'])
        empty_df.to_excel(excel_path, index=False)
        return True
    return False

def load_from_cache(cache_path, progress):
    progress.update(30, "Opening cache")
    with open(cache_path, 'rb') as f:
        progress.update(60, "Reading cache")
        df = pickle.load(f)
    progress.update(90, "Verifying cache")
    progress.update(100, f"Loaded {len(df)} tires from cache")
    return df

def load_data(progress):
    excel_path = get_excel_path()
    cache_path = get_cache_path()
    hash_path = get_hash_path()
    
    # Create Excel if not exists
    if not os.path.exists(excel_path):
        progress.update(50, "Creating new database")
        create_empty_excel()
        progress.update(100, "Created empty database")
        time.sleep(0.5)
        return pd.DataFrame(columns=['ID', 'BRAND', 'WIDTH', 'PROFILE', 'DIAMETER', 
                                      'SEASON', 'CONDITION', 'SECTOR', 'FLOOR', 'QUANTITY'])
    
    # Check cache
    progress.update(12, "Checking cache")
    current_hash = get_file_hash(excel_path)
    
    if os.path.exists(cache_path) and os.path.exists(hash_path) and current_hash:
        with open(hash_path, 'r') as f:
            cached_hash = f.read()
        
        if cached_hash == current_hash:
            progress.update(15, "Cache valid")
            df = load_from_cache(cache_path, progress)
            return df
    
    # Load from Excel
    progress.update(15, "Cache invalid, loading Excel")
    df = load_excel_with_progress(excel_path, progress)
    
    # Save to cache
    progress.update(99, "Saving to cache")
    with open(cache_path, 'wb') as f:
        pickle.dump(df, f)
    with open(hash_path, 'w') as f:
        f.write(str(current_hash))
    
    return df

def save_data(df):
    excel_path = get_excel_path()
    cache_path = get_cache_path()
    hash_path = get_hash_path()
    
    print("\nSaving data...")
    
    try:
        df.to_excel(excel_path, index=False, engine='openpyxl')
        print("  Saved to Excel")
    except Exception as e:
        print(f"  ERROR: {e}")
        return False
    
    try:
        if os.path.exists(cache_path):
            os.remove(cache_path)
        if os.path.exists(hash_path):
            os.remove(hash_path)
        print("  Cache cleared")
    except:
        pass
    
    gc.collect()
    return True

def safe_int(value, default=0):
    try:
        if value is None or value == '' or value == 'nan':
            return default
        if isinstance(value, float):
            return int(value) if not pd.isna(value) else default
        if isinstance(value, str):
            value = value.strip().replace('.0', '').replace(',', '').replace(' ', '')
            if value == '':
                return default
            return int(float(value))
        return int(value)
    except:
        return default

def get_next_id(df):
    if len(df) == 0:
        return 1
    return int(df['ID'].max()) + 1

def add_tire(df):
    print("\n" + "="*60)
    print("ADD NEW TIRE")
    print("="*60)
    
    print("\nEnter tire information:")
    print("-"*40)
    
    try:
        brand = input("Brand: ").strip()
        if not brand:
            print("Brand cannot be empty!")
            return df
        
        width = input("Width: ").strip()
        profile = input("Profile: ").strip()
        diameter = input("Diameter: ").strip()
        
        if not width or not profile or not diameter:
            print("Size cannot be empty!")
            return df
        
        width = safe_int(width)
        profile = safe_int(profile)
        diameter = safe_int(diameter)
        
        season = input("Season (Summer/Winter/AllSeason): ").strip()
        if not season:
            season = "Summer"
        
        condition = input("Condition (New/Used): ").strip()
        if not condition:
            condition = "New"
        
        sector = input("Sector (A, B, C, etc.): ").strip().upper()
        if not sector:
            sector = "A"
        
        floor = input("Floor: ").strip()
        floor = safe_int(floor)
        
        quantity = input("Quantity: ").strip()
        quantity = safe_int(quantity, 1)
        
        new_id = get_next_id(df)
        new_tire = {
            'ID': new_id,
            'BRAND': brand,
            'WIDTH': width,
            'PROFILE': profile,
            'DIAMETER': diameter,
            'SEASON': season,
            'CONDITION': condition,
            'SECTOR': sector,
            'FLOOR': floor,
            'QUANTITY': quantity
        }
        
        df = pd.concat([df, pd.DataFrame([new_tire])], ignore_index=True)
        df = sort_tires(df)
        save_data(df)
        
        print("\n" + "="*40)
        print("SUCCESS! Tire added!")
        print(f"Brand: {brand}")
        print(f"Size: {width}/{profile}/R{diameter}")
        print(f"Location: Sector {sector}, Floor {floor}")
        print(f"Quantity: {quantity} pcs")
        print("="*40)
        
    except Exception as e:
        print(f"Error: {e}")
    
    input("\nPress Enter to continue...")
    return df

def search_by_size(df, width, profile, diameter):
    width = safe_int(width)
    profile = safe_int(profile)
    diameter = safe_int(diameter)
    result = df[(df['WIDTH'] == width) & (df['PROFILE'] == profile) & (df['DIAMETER'] == diameter)]
    return result

def sell_tire_by_id(df):
    print("\n" + "="*60)
    print("SELL TIRES")
    print("="*60)
    
    while True:
        print("\nEnter tire ID (or 0 to exit):")
        
        id_input = input("ID: ").strip()
        
        if id_input == "0":
            return df
        
        if not id_input.isdigit():
            print("Invalid ID!")
            continue
        
        tire_id = int(id_input)
        tire_index = df[df['ID'] == tire_id].index
        
        if len(tire_index) == 0:
            print(f"No tire with ID: {tire_id}")
            continue
        
        idx = tire_index[0]
        tire = df.loc[idx]
        current_qty = safe_int(tire['QUANTITY'])
        
        print("\n" + "-"*50)
        print(f"ID: {tire['ID']}")
        print(f"Brand: {tire['BRAND']}")
        print(f"Size: {tire['WIDTH']}/{tire['PROFILE']}/R{tire['DIAMETER']}")
        print(f"Location: Sector {tire['SECTOR']}, Floor {tire['FLOOR']}")
        print(f"Quantity: {current_qty} pcs")
        print("-"*50)
        
        qty_input = input("\nHow many to sell (Enter = all): ")
        
        try:
            if qty_input == "":
                df = df.drop(idx)
                print("Position removed!")
            else:
                qty = safe_int(qty_input)
                if qty <= 0:
                    print("Invalid quantity!")
                    continue
                elif qty >= current_qty:
                    df = df.drop(idx)
                    print("Position removed!")
                else:
                    df.loc[idx, 'QUANTITY'] = current_qty - qty
                    print(f"Sold {qty}. Remaining: {df.loc[idx, 'QUANTITY']}")
            
            df = sort_tires(df)
            save_data(df)
            break
            
        except Exception as e:
            print(f"Error: {e}")
    
    return df

def display_results_table(results):
    if len(results) == 0:
        print("\nNo tires found!")
        return False
    
    results = sort_tires(results)
    
    print("\n" + "="*100)
    print(f"FOUND: {len(results)} POSITIONS")
    print("="*100)
    print(f"{'ID':<5} {'BRAND':<12} {'SIZE':<12} {'SEASON':<10} {'CONDITION':<12} {'SECT':<5} {'FLR':<5} {'QTY':<5}")
    print("-"*100)
    
    total_qty = 0
    for _, row in results.iterrows():
        size = f"{row['WIDTH']}/{row['PROFILE']}/R{row['DIAMETER']}"
        print(f"{row['ID']:<5} {row['BRAND']:<12} {size:<12} {row['SEASON']:<10} {row['CONDITION']:<12} {row['SECTOR']:<5} {row['FLOOR']:<5} {row['QUANTITY']:<5}")
        total_qty += row['QUANTITY']
    
    print("="*100)
    print(f"TOTAL: {total_qty} pcs")
    print("="*100)
    return True

def display_all_table(df):
    if len(df) == 0:
        print("\nNo tires in stock!")
        input("Press Enter...")
        return
    
    print("\n" + "="*100)
    print("ALL TIRES IN STOCK")
    print("="*100)
    print(f"{'ID':<5} {'BRAND':<12} {'SIZE':<12} {'SEASON':<10} {'CONDITION':<12} {'SECT':<5} {'FLR':<5} {'QTY':<5}")
    print("-"*100)
    
    total_qty = 0
    for _, row in df.iterrows():
        size = f"{row['WIDTH']}/{row['PROFILE']}/R{row['DIAMETER']}"
        print(f"{row['ID']:<5} {row['BRAND']:<12} {size:<12} {row['SEASON']:<10} {row['CONDITION']:<12} {row['SECTOR']:<5} {row['FLOOR']:<5} {row['QUANTITY']:<5}")
        total_qty += row['QUANTITY']
    
    print("="*100)
    print(f"TOTAL: {total_qty} pcs")
    print("="*100)
    input("Press Enter...")

def search_by_brand(df, brand):
    result = df[df['BRAND'].str.contains(brand, case=False, na=False)]
    return sort_tires(result)

def search_by_size_menu(df):
    print("\n" + "="*60)
    print("SEARCH BY SIZE")
    print("="*60)
    print("\nEnter size (e.g., 205/55/16) or 0 to exit:")
    
    size_input = input("Size: ").strip()
    
    if size_input == "0":
        return
    
    if "/" in size_input:
        parts = size_input.split("/")
        if len(parts) == 3:
            width = parts[0]
            profile = parts[1]
            diameter = parts[2].replace("R", "")
            results = search_by_size(df, width, profile, diameter)
            display_results_table(results)
            if len(results) > 0:
                input("Press Enter...")
        else:
            print("Wrong format! Use: 205/55/16")
            input("Press Enter...")
    else:
        print("Wrong format! Use: 205/55/16")
        input("Press Enter...")

def search_by_brand_menu(df):
    print("\n" + "="*60)
    print("SEARCH BY BRAND")
    print("="*60)
    print("\nEnter brand name or 0 to exit:")
    
    brand = input("Brand: ").strip()
    
    if brand == "0":
        return
    
    if brand:
        results = search_by_brand(df, brand)
        if len(results) > 0:
            display_results_table(results)
            input("Press Enter...")
        else:
            print(f"Brand '{brand}' not found!")
            input("Press Enter...")

def show_all(df):
    display_all_table(df)

def main():
    # Animation is already running!
    progress.update(2, "Loading configuration")
    time.sleep(0.1)
    
    # Load data
    df = load_data(progress)
    
    # Finish animation
    progress.finish(f"Ready! {len(df)} tires in stock")
    
    print("\n" + "="*50)
    print("TIRE STOCK MANAGEMENT SYSTEM")
    print("="*50)
    
    while True:
        print("\n" + "="*50)
        print("1. Search by size")
        print("2. Search by brand")
        print("3. Show all tires")
        print("4. Add new tire")
        print("5. Sell tire (by ID)")
        print("6. Exit")
        print("="*50)
        
        choice = input("Select option (1-6): ")
        
        if choice == "1":
            search_by_size_menu(df)
        elif choice == "2":
            search_by_brand_menu(df)
        elif choice == "3":
            show_all(df)
        elif choice == "4":
            df = add_tire(df)
        elif choice == "5":
            df = sell_tire_by_id(df)
        elif choice == "6":
            print("\nGoodbye!")
            break
        else:
            print("Invalid choice!")

if __name__ == "__main__":
    main()
