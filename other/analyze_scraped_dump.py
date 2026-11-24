import json

with open(r'Invoice_Indexer/masters/scraped_extjs_dump.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

stores = data.get('stores', {}).get('stores', [])
active_stores = [s for s in stores if s.get('count', 0) > 0]
print('Active stores with data (>0):')
for s in active_stores:
    print(' ', s['id'], ':', s['count'], 'records')

props = data.get('stores', {}).get('properties', [])
vendors = data.get('stores', {}).get('vendors', [])
print(f'\nExtJS Scraped Properties in memory: {len(props)}')
if props:
    print('Sample property:', props[0])

print(f'ExtJS Scraped Vendors in memory: {len(vendors)}')
if vendors:
    print('Sample vendor:', vendors[0])

captured = data.get('captured_api_payloads', [])
print(f'\nCaptured API endpoints: {len(captured)}')
for c in captured:
    url = c.get('url', '')
    print('  URL:', url)
    d = c.get('data')
    if isinstance(d, dict):
        if 'data' in d and isinstance(d['data'], list):
            print(f'    Count: {len(d["data"])}')
            if d['data']:
                print('    Sample:', d['data'][0])
        else:
            print('    Keys:', list(d.keys())[:5])
    elif isinstance(d, list):
        print(f'    List items: {len(d)}')
        if d:
            print('    Sample:', d[0])
