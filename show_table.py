import json

with open('bid_data.json') as f:
    data = json.load(f)

data = [s for s in data if s['name'].strip()]
data.sort(key=lambda x: x['name'])

header = f"{'Name':<30} {'Wkly Hrs':>8}  {'Sun':>5} {'Mon':>5} {'Tue':>5} {'Wed':>5} {'Thu':>5} {'Fri':>5} {'Sat':>5}"
print(header)
print('-' * len(header))

for s in data:
    days = []
    for d in ['Sun', 'Mon', 'Tue', 'Wed', 'Thur', 'Fri', 'Sat']:
        info = s['days'].get(d, {})
        h = info.get('hrs', '')
        if info.get('time') == 'OFF':
            h = 'OFF'
        elif not info.get('time'):
            h = '--'
        days.append(h)
    name = s['name']
    wh = s['wkly_hrs']
    print(f"{name:<30} {wh:>8}  {days[0]:>5} {days[1]:>5} {days[2]:>5} {days[3]:>5} {days[4]:>5} {days[5]:>5} {days[6]:>5}")
