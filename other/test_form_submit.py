import urllib.request
import re

url = 'https://docs.google.com/forms/d/e/1FAIpQLScIqyCFFiE4XSYCxGsMWdpkzQkidzsnX8Yzk-4b4CeOpo7mWQ/viewform'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
with urllib.request.urlopen(req) as resp:
    html = resp.read().decode('utf-8')
    action = re.search(r'<form\s+action="([^"]+)"', html)
    print('Form action:', action.group(1) if action else 'Not found')
