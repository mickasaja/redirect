#!/usr/bin/env python3
"""Build the wedding RSVP guest list from Google Sheets."""
import urllib.request, csv, io, json

url = "https://docs.google.com/spreadsheets/d/1IbeHKa3JtJIEE4uCcDIZBDWAmiS85KG_CuGeXwIONkA/export?format=csv&gid=0"
resp = urllib.request.urlopen(url)
content = resp.read().decode('utf-8')
reader = csv.reader(io.StringIO(content))
next(reader)

guests = {}
countries = {}

for row in reader:
    if len(row) < 7: continue
    first = row[2].strip()
    status = (row[3] or '').strip()
    count_str = (row[4] or '0').strip()
    cc = (row[5] or '').strip()
    phone = (row[6] or '').strip()
    
    if status.lower() != 'main' or not phone or not cc: continue
    
    # Skip bad entries
    try: int(cc)
    except: continue
    if cc == 'VALUE!': continue
    
    key = f"+{cc}{phone}"
    cnt = int(count_str) if count_str.isdigit() else 0
    
    guests[key] = {"first": first or "Guest", "guests": cnt, "cc": cc}

print(f"// AUTO-GENERATED from Google Sheets — {len(guests)} guests")
print("var GUESTS={")
for k, v in sorted(guests.items()):
    print(f'  "{k}":{{first:"{v["first"]}",guests:{v["guests"]}}},')
print("};")
print()

# Also generate country codes
cc_set = sorted(set(v['cc'] for v in guests.values()), key=lambda x: int(x) if x.isdigit() else 999)
for cc in cc_set:
    print(f"// CC {cc}: {sum(1 for v in guests.values() if v['cc']==cc)} guests")