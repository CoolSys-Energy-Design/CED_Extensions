# generic_triggers.py
# Project-AGNOSTIC trigger discovery for PF yaml profiles.
# Inputs: Revit spaces (rooms) + CAD blocks (gym/spa equipment positions).
# Output: dict {profile_name: [(anchor_x, anchor_y, anchor_rotation_deg), ...]}.
#
# Works for ANY PF project as long as room names and CAD layers follow PF conventions.

import math
import clr
from Autodesk.Revit import DB

_NAME = clr.GetClrType(DB.Element).GetProperty('Name')

# Profile name -> room-name-pattern-based discovery rules
# Each rule: pattern strings to match in room Name, optional excludes, optional rotation.
PROFILE_ROOM_RULES = {
    # Spa rooms (each room → 1 anchor at room center, fixed wall direction)
    'BCS_Tanning Red Wave':        {'patterns': ['RED WAVE'],          'exclude': [],                      'anchor_rot': 0,   'multi': 1},
    'BCS_Tanning Bed 42.4':        {'patterns': ['TANNING'],           'exclude': ['SPRAY','RED WAVE','HYBRID'], 'anchor_rot': 0,   'multi': 1},
    'BCS_Hybrid Tanning Booth':    {'patterns': ['HYBRID'],            'exclude': [],                      'anchor_rot': 0,   'multi': 1},
    'BCS_Cryolounge':              {'patterns': ['CRYO'],              'exclude': [],                      'anchor_rot': 270, 'multi': 3},
    'BCS_Hydro Lounge':            {'patterns': ['HYDRO'],             'exclude': [],                      'anchor_rot': 90,  'multi': 4},
    'SAUNA REDZONE':               {'patterns': ['SAUNA'],             'exclude': [],                      'anchor_rot': 270, 'multi': 1},
    'Spray Tanning':               {'patterns': ['SPRAY TAN'],         'exclude': [],                      'anchor_rot': 90,  'multi': 1},
    # Building-service rooms
    'IT RACKS':                    {'patterns': ['IT ROOM'],           'exclude': [],                      'anchor_rot': 270, 'multi': 1},
    'PF_Plan Computer':            {'patterns': ['IT ROOM'],           'exclude': [],                      'anchor_rot': 270, 'multi': 1},
    'Room Name 101 RECEPTION':     {'patterns': ['RECEPTION'],         'exclude': [],                      'anchor_rot': 90,  'multi': 1},
    'Room Name 101A ROOM MECHANICAL': {'patterns': ['MECHANICAL'],     'exclude': [],                      'anchor_rot': 0,   'multi': 1},
    'Room Name 105 LOCKER ROOM MEN\'S': {'patterns': ['MENS LOCKER'],  'exclude': [],                      'anchor_rot': 0,   'multi': 1},
    'Room Name 104 LOCKER ROOM WOMEN\'S': {'patterns': ['WOMENS LOCKER'], 'exclude': [],                   'anchor_rot': 0,   'multi': 1},
    'PF_Plan Rinnai HWH':          {'patterns': ['STORAGE ROOM','MECHANICAL'], 'exclude': [],              'anchor_rot': 0,   'multi': 1},
    # Restroom hand dryers - one per toilet room
    'PF_Plan Hand Dryer':          {'patterns': ['TOILET'],            'exclude': [],                      'anchor_rot': 0,   'multi': 1},
    # Sloan trough sinks - under STORAGE rooms (105D, 104D)
    'PF_Plan Sloan Trough Sink':   {'patterns': ['STORAGE'],           'exclude': ['STORAGE ROOM'],        'anchor_rot': 0,   'multi': 1, 'y_offset_from_room': -1.75},
    # Soap dispensers - same location as sinks
    'PF_Plan Soap Dispenser':      {'patterns': ['STORAGE'],           'exclude': ['STORAGE ROOM'],        'anchor_rot': 0,   'multi': 1, 'y_offset_from_room': -0.25},
    # Vending machines - vestibule
    'vending machine':             {'patterns': ['VESTIBULE'],         'exclude': [],                      'anchor_rot': 0,   'multi': 1},
    'PF_Plan Vend Mach':           {'patterns': ['VESTIBULE'],         'exclude': [],                      'anchor_rot': 0,   'multi': 1},
    # BCS chairs and angels (multiple per BCS room - placed evenly)
    'BCS_Beauty Angel':            {'patterns': ['BLACK CARD'],        'exclude': [],                      'anchor_rot': 180, 'multi': 4, 'x_spread': 30},
    'BCS_SmarteCarte Massage Chair': {'patterns': ['BLACK CARD'],      'exclude': [],                      'anchor_rot': 0,   'multi': 2, 'x_spread': 16},
    'treatment room chair':        {'patterns': ['BLACK CARD'],        'exclude': [],                      'anchor_rot': 0,   'multi': 1},
    # Top-of-strength TVs (43") - placed near free weights / strength
    '43 tv':                       {'patterns': ['STRENGTH','FREE WEIGHTS'], 'exclude': [],                'anchor_rot': 180, 'multi': 5, 'wall_snap': 'north', 'x_spread': 14},
    # Locker room TVs (from CAD blocks, handled separately)
    'PF_Plan TV':                  {'patterns': [],                    'exclude': [],                      'anchor_rot': 0,   'multi': 0},  # use CAD blocks
    # Drinking fountain - in corridor between bathrooms
    'PF_Plan DF':                  {'patterns': [],                    'exclude': [],                      'anchor_rot': 0,   'multi': 0},  # use PLUMB FIX block
}


def get_rooms(doc, view):
    """Return list of {'name', 'num', 'xy'} for all spaces in view."""
    rooms = []
    for sp in DB.FilteredElementCollector(doc, view.Id).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType():
        try:
            nm = sp.LookupParameter('Name'); name = nm.AsString() if nm else ''
            nu = sp.LookupParameter('Number'); num = nu.AsString() if nu else ''
            loc = sp.Location
            if not isinstance(loc, DB.LocationPoint): continue
            rooms.append({'name': name or '', 'num': num or '', 'xy': (loc.Point.X, loc.Point.Y)})
        except: pass
    return rooms


def match_rooms(rooms, patterns, excludes):
    """Return rooms whose name contains any pattern AND none of the excludes."""
    out = []
    for r in rooms:
        n = (r['name'] or '').upper()
        if any(p.upper() in n for p in patterns) and not any(x.upper() in n for x in excludes):
            out.append(r)
    return out


def discover_triggers(doc, view, cardio_pts, plumb_fix_blocks, locker_tv_blocks):
    """Build the full {profile_name: [(x,y,rot_deg)]} dict for any PF project.

    Args:
        cardio_pts: list of (x, y) - cardio north-edge centers (for treadmill + TV truss)
        plumb_fix_blocks: list of dicts with 'world_xy' - drinking fountain / hand dryer candidates
        locker_tv_blocks: list of dicts with 'world_xy' and 'rotation_deg' - TV blocks in lockers
    """
    rooms = get_rooms(doc, view)
    triggers = {}

    # --- CARDIO equipment (Treadmill) ---
    if cardio_pts:
        triggers['Treadmill (T5X-5PL-PF)'] = [(x, y, 0) for x, y in cardio_pts]
        # TV truss anchored above the cardio area
        if len(cardio_pts) >= 2:
            xs = sorted([x for x, _ in cardio_pts])
            ys = [y for _, y in cardio_pts]
            truss_y = max(ys) + 6.0   # 6 ft north of topmost cardio row
            # 7 trusses evenly spaced across cardio X range
            xmin, xmax = xs[0], xs[-1]
            n_truss = 7
            truss_xs = [xmin + i*(xmax-xmin)/(n_truss-1) for i in range(n_truss)]
            triggers['TV TRUSS'] = [(x, truss_y, 0) for x in truss_xs]

    # --- Locker room TVs from CAD A-N-TELEVISION blocks ---
    if locker_tv_blocks:
        triggers['PF_Plan TV'] = [(b['world_xy'][0], b['world_xy'][1], b.get('rotation_deg', 0)) for b in locker_tv_blocks]

    # --- Drinking fountain from corridor PLUMB FIX block (Y > -280 typically) ---
    if plumb_fix_blocks:
        df_candidates = [b for b in plumb_fix_blocks if b['world_xy'][1] > -280]
        if df_candidates:
            triggers['PF_Plan DF'] = [(b['world_xy'][0], b['world_xy'][1], 0) for b in df_candidates[:1]]

    # --- Room-name-pattern profiles ---
    for prof_name, rule in PROFILE_ROOM_RULES.items():
        patterns = rule.get('patterns', [])
        if not patterns: continue
        excludes = rule.get('exclude', [])
        matched = match_rooms(rooms, patterns, excludes)
        if not matched: continue
        anchor_rot = rule.get('anchor_rot', 0)
        multi = rule.get('multi', 1)
        x_spread = rule.get('x_spread', 0)
        y_offset = rule.get('y_offset_from_room', 0)
        anchors = []
        for r in matched:
            cx, cy = r['xy']
            cy += y_offset
            if multi == 1:
                anchors.append((cx, cy, anchor_rot))
            elif multi > 1 and x_spread > 0:
                # Multiple instances spread along X
                spacing = x_spread / 12.0   # convert inches to feet
                for i in range(multi):
                    off = (i - (multi-1)/2.0) * spacing
                    anchors.append((cx + off, cy, anchor_rot))
            else:
                # Multiple beds at same anchor (yaml offsets will spread them later)
                # For Cryolounge/Hydromassage, use the bed positions inside the room
                # Default: stack multiple anchors at same XY (yaml offsets will handle)
                for i in range(multi):
                    # Stagger Y by 7ft for typical bed spacing (hydro lounge)
                    off_y = (i - (multi-1)/2.0) * 7.0
                    anchors.append((cx, cy + off_y, anchor_rot))
        triggers[prof_name] = anchors

    return triggers
