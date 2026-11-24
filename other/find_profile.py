import os
import json

user_data = r"C:\Users\Meet Patel\AppData\Local\Google\Chrome\User Data"
profiles = [d for d in os.listdir(user_data) if os.path.isdir(os.path.join(user_data, d)) and ('Profile' in d or d == 'Default')]
print('Chrome profile directories:', profiles)

for p in profiles:
    pref_path = os.path.join(user_data, p, 'Preferences')
    if os.path.exists(pref_path):
        try:
            with open(pref_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                email = data.get('account_info', [{}])[0].get('email') if data.get('account_info') else None
                name = data.get('profile', {}).get('name')
                print(f'Profile "{p}": name="{name}", email="{email}"')
        except Exception as e:
            print(f'Profile "{p}" error:', e)
