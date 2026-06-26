#!/usr/bin/env python3
"""Build the wedding RSVP guest list from Google Sheets — writes directly to file."""
import urllib.request, csv, io

url = "https://docs.google.com/spreadsheets/d/1IbeHKa3JtJIEE4uCcDIZBDWAmiS85KG_CuGeXwIONkA/export?format=csv&gid=0"
resp = urllib.request.urlopen(url)
content = resp.read().decode('utf-8')
reader = csv.reader(io.StringIO(content))
next(reader)

lines = []
lines.append('// AUTO-GENERATED from Google Sheets')
lines.append('var GUESTS={')

count = 0
for row in reader:
    if len(row) < 7: continue
    first = row[2].strip()
    status = (row[3] or '').strip()
    count_str = (row[4] or '0').strip()
    cc = (row[5] or '').strip()
    phone = (row[6] or '').strip()
    
    if status.lower() != 'main' or not phone or not cc: continue
    try: int(cc)
    except: continue
    if cc == 'VALUE!': continue
    
    key = f"+{cc}{phone}"
    cnt = int(count_str) if count_str.isdigit() else 0
    name = first or 'Guest'
    
    lines.append(f'  "{key}":{{first:"{name}",guests:{cnt}}},')
    count += 1

lines.append('};')

# Write to file directly (avoid terminal redaction)
with open('guest_list.js', 'w') as f:
    f.write('\n'.join(lines))

print(f'Wrote {count} guests to guest_list.js')
