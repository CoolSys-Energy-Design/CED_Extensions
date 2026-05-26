# charlotte_triggers.py
# Defines trigger positions per yaml profile for the Charlotte project.
# Each entry: profile_name -> list of (anchor_x_ft, anchor_y_ft, anchor_rotation_deg)
# The yaml profile's linked elements are then placed at: anchor + rotated(offset)
#
# Notes for trigger anchor rotation convention:
#   anchor_rotation_deg = 0   -> equipment "faces north" (+Y) in world
#   anchor_rotation_deg = 90  -> equipment "faces west" (-X)
#   anchor_rotation_deg = 180 -> faces south
#   anchor_rotation_deg = 270 -> faces east
# This rotation is applied to the yaml offset (which is in equipment-local coords).
#
# When the yaml offset has rotation_deg field, that becomes the FIXTURE's own rotation
# (e.g. how the recep is oriented) AFTER being rotated by the anchor rotation.

import math

def get_charlotte_triggers(cardio_pts):
    """Returns dict {profile_name: [(anchor_x, anchor_y, anchor_rot_deg), ...]}.
    cardio_pts: list of (x, y) for cardio block north-edge centers (passed in)."""
    triggers = {}

    # ===== CARDIO (Treadmill) =====
    # Anchor at each cardio block's north-edge center, facing north (rot=0)
    triggers['Treadmill (T5X-5PL-PF)'] = [(x, y, 0) for x, y in cardio_pts]

    # ===== TV TRUSS =====
    # The truss runs above the cardio area E-W. Place ~7 TVs evenly spaced along it.
    # Based on cardio X range ~106..205 and reference position around Y=-225 (above N cardio row)
    tv_truss_xs = [120, 135, 150, 165, 180, 195, 210]
    triggers['TV TRUSS'] = [(x, -224.5, 0) for x in tv_truss_xs]

    # ===== TANNING ROOMS - per-room profiles =====
    # 103B RED WAVE
    triggers['BCS_Tanning Red Wave'] = [(214.60, -299.83, 0)]
    # 103C, 103E, 103K, 103L - Tanning Bed 42.4 (4 rooms)
    triggers['BCS_Tanning Bed 42.4'] = [
        (224.37, -300.07, 0),  # 103C - south wall
        (244.08, -300.26, 0),  # 103E - south wall
        (230.68, -284.79, 90), # 103K - east wall (faces west)
        (221.13, -285.08, 270),# 103L - west wall (faces east)
    ]
    # 103D, 103F - HYBRID TANNING BOOTH
    triggers['BCS_Hybrid Tanning Booth'] = [
        (234.25, -299.85, 0),  # 103D
        (254.22, -300.01, 0),  # 103F
    ]
    # 103G CRYOLOUNGE - 3 beds on east wall
    triggers['BCS_Cryolounge'] = [
        (269.4, -283.0, 270),  # bed 1
        (269.4, -289.0, 270),  # bed 2
        (269.4, -295.0, 270),  # bed 3
    ]
    # 103A HYDROMASSAGE - 4 beds on west wall
    triggers['BCS_Hydro Lounge'] = [
        (197.5, -278.5, 90),
        (197.5, -285.7, 90),
        (197.5, -292.8, 90),
        (197.5, -300.0, 90),
    ]
    # 103H SAUNA - east wall
    triggers['SAUNA REDZONE'] = [(254.5, -285.73, 270)]
    # 103J SPRAY TANNING - west wall
    triggers['Spray Tanning'] = [(235.5, -286.03, 90)]
    # BCS_Beauty Angel - in BCS area, 4 units along north wall
    triggers['BCS_Beauty Angel'] = [(x, -270.5, 180) for x in [222, 232, 242, 252]]
    # BCS_SmarteCarte Massage Chair - 2 floor outlets in BCS
    triggers['BCS_SmarteCarte Massage Chair'] = [(228, -276, 0), (244, -276, 0)]
    triggers['treatment room chair'] = [(218, -274, 0)]

    # ===== TV CATEGORIES =====
    # 43 tv - 5 TVs at top of strength area (already determined as wall snap)
    triggers['43 tv'] = [(x, -190.11, 180) for x in [167.36, 170.78, 174.20, 177.61, 180.61]]
    # PF_Plan TV - locker room TVs from A-N-TELEVISION blocks
    triggers['PF_Plan TV'] = [(100.09, -283.96, -45), (182.5, -283.96, 45)]

    # ===== IT ROOM =====
    triggers['IT RACKS'] = [(269, -268.63, 270)]   # east wall, facing west

    # ===== RESTROOMS / LOCKERS / RECEPTION =====
    # PF_Plan Sloan Trough Sink - 2 locations (under STORAGE rooms per LEARNINGS)
    triggers['PF_Plan Sloan Trough Sink'] = [
        (136.12, -285.0, 0),   # Men's sink (under 105D STORAGE)
        (146.29, -285.0, 0),   # Women's sink (under 104D STORAGE)
    ]
    # PF_Plan Hand Dryer - restroom hand dryers
    triggers['PF_Plan Hand Dryer'] = [
        (151.0, -296.0, 0),  # 104A TOILET
        (130.7, -296.1, 0),  # 105A TOILET
    ]
    # PF_Plan DF - drinking fountain at the corridor PLUMB FIX position
    triggers['PF_Plan DF'] = [(198.1, -272.1, 0)]
    # PF_Plan Soap Dispenser - one per locker room sink
    triggers['PF_Plan Soap Dispenser'] = [(136.12, -283.5, 0), (146.29, -283.5, 0)]
    # PF_Plan Rinnai HWH - water heater (typically in mech room or storage)
    triggers['PF_Plan Rinnai HWH'] = [(190.0, -294.2, 0)]  # near STORAGE 101D

    # ===== VENDING / VESTIBULE / RECEPTION =====
    triggers['vending machine'] = [(263.0, -244.81, 0)]      # vestibule
    triggers['PF_Plan Vend Mach'] = [(263.0, -244.81, 0)]    # alias
    # Room Name 101 RECEPTION - 6 fixtures along reception counter
    triggers['Room Name 101 RECEPTION'] = [(256.29, -250.48, 90)]  # east wall area
    # Room Name 101A ROOM MECHANICAL - 3 fixtures in mech room
    triggers['Room Name 101A ROOM MECHANICAL'] = [(260.0, -262.0, 0)]
    # Locker room TVs etc.
    triggers["Room Name 105 LOCKER ROOM MEN'S"] = [(115.45, -276.25, 0)]
    triggers["Room Name 104 LOCKER ROOM WOMEN'S"] = [(168.0, -277.2, 0)]
    triggers['Room Name 104 LOCKER ROOM'] = []  # generic, skip in favor of women's-specific

    # ===== MISC =====
    triggers['PF_Plan Computer'] = [(263.76, -266.0, 270)]   # IT room
    triggers['it'] = []  # skip duplicate
    triggers['_skip PF_Plan TV Truss'] = []  # explicitly skipped

    return triggers
