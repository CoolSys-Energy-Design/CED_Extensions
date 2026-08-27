# Build injection/manifest.json from recreation.json — the machine-readable placement order.
import json, os, collections
R = json.load(open('analysis/recreation.json'))
os.makedirs('injection', exist_ok=True)

manifest = {
  'meta': {
    'project': 'Planet Fitness — CHARLOTTE (CENTRAL), NC 14202-25',
    'view': 'E101 - Power Plan',
    'level': 'L1 - Finished Floor',
    'units': 'feet, project internal coordinates (same origin as the CAD background X_BG.dwg)',
    'generator': 'pf_e101 reconstruct.py — CAD-driven rules derived from PF2_Training (26 stores)',
  },
  'devices': [],
  'panels': R['panels'],
  'circuits': [],
  'keynotes': R['keynotes'],
}
for d in R['devices']:
    manifest['devices'].append({
        'family': d['fam'], 'type': d['typ'],
        'x': d['x'], 'y': d['y'], 'z': 0.0,
        'rotation_deg': d['rot'], 'elevation_from_level': d['elev'],
        'load_name': d['load'], 'va': d['va'],
        'panel': d['panel'], 'circuit': d['ckt'],
        'keynote': d['kn'], 'driver': d['driver'], 'why': d['why'],
    })
for c in R['circuits']:
    manifest['circuits'].append({
        'panel': c['panel'], 'circuit': c['ckt'], 'load_name': c['load'],
        'va': c['va'], 'volts': c['volts'], 'poles': c['poles'], 'rating_a': c['rating'],
        'member_device_indices': c['members'],
    })
json.dump(manifest, open('injection/manifest.json','w'), indent=1)
ft = collections.Counter((d['family'], d['type']) for d in manifest['devices'])
pf = collections.Counter((p['fam'], p['typ']) for p in manifest['panels'])
json.dump({'device_types': [[f,t,n] for (f,t),n in ft.most_common()],
           'panel_types': [[f,t,n] for (f,t),n in pf.most_common()]},
          open('injection/famtypes.json','w'), indent=1)
print('manifest devices:', len(manifest['devices']), 'circuits:', len(manifest['circuits']),
      'panels:', len(manifest['panels']), 'keynotes:', len(manifest['keynotes']))
print('distinct device famtypes:', len(ft))
