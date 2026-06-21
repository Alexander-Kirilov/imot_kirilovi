"""
patch_bulk_import.py
--------------------
Еднократен скрипт — коригира Bulk_Import колоната в съществуващия parquet.

Логика: ако на дадена First_Seen_Date са добавени > BULK_THRESHOLD обяви,
те се маркират като Bulk_Import=True (добавени при масов/начален run).

Пусни ВЕДНЪЖ преди да пуснеш обновения скрапер.
"""

import pandas as pd
from pathlib import Path
import shutil
from datetime import datetime

HISTORY_FILE = "all_listings_history.parquet"

# Threshold за bulk: дни с повече от толкова нови обяви = bulk run.
# При нормален ежедневен скрапинг реалистично се появяват 1-5 нови на ден.
# Стойност 10 хваща очевидните масови imports (40, 58, 63 обяви),
# без да засяга нормални дни.
BULK_THRESHOLD = 10

# ── Зареждане ──────────────────────────────────────────────────────────────────
p = Path(HISTORY_FILE)
if not p.exists():
    print(f"❌ Файлът {HISTORY_FILE} не е намерен!")
    exit(1)

df = pd.read_parquet(HISTORY_FILE)
print(f"Заредени {len(df)} реда от {HISTORY_FILE}")

# ── Backup ─────────────────────────────────────────────────────────────────────
backup_name = f"all_listings_history_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.parquet"
shutil.copy2(HISTORY_FILE, backup_name)
print(f"✓ Backup: {backup_name}")

# ── Анализ по дата ─────────────────────────────────────────────────────────────
counts = df.groupby('First_Seen_Date').size()
bulk_dates = counts[counts > BULK_THRESHOLD].index.tolist()

print(f"\nОбяви по First_Seen_Date:")
for date, cnt in counts.sort_index().items():
    marker = f"← BULK (>{BULK_THRESHOLD})" if date in bulk_dates else ""
    print(f"  {date}: {cnt:4d} {marker}")

# ── Patch ──────────────────────────────────────────────────────────────────────
before_bulk = df['Bulk_Import'].fillna(False).sum()
df['Bulk_Import'] = df['First_Seen_Date'].isin(bulk_dates)
after_bulk = df['Bulk_Import'].sum()

print(f"\nBulk_Import преди: {int(before_bulk)}")
print(f"Bulk_Import след:  {int(after_bulk)}")
print(f"Променени редове:  {int(after_bulk - before_bulk)}")

# ── Запис ─────────────────────────────────────────────────────────────────────
df.to_parquet(HISTORY_FILE, index=False)
print(f"\n✓ Записано в {HISTORY_FILE}")

# ── Проверка: таб Нови ─────────────────────────────────────────────────────────
from datetime import timedelta
cutoff = (datetime.now() - timedelta(days=5)).strftime("%Y-%m-%d")
df_active = df[~df['Sold'].fillna(False)]
mask_recent = df_active['First_Seen_Date'].fillna('').astype(str) >= cutoff
mask_not_bulk = ~df_active['Bulk_Import'].fillna(False)
df_new_tab = df_active[mask_recent & mask_not_bulk]

print(f"\n=== Симулация таб 'Нови' (cutoff={cutoff}) ===")
print(f"Показвани: {len(df_new_tab)} обяви")
if not df_new_tab.empty:
    print(df_new_tab[['First_Seen_Date', 'Location', 'Price_EUR']].to_string())
