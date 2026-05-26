# Lindenhurst Power Plan — QAQC vs Charlotte E101

**Reference:** Charlotte (Central), NC — view "E101 - Power Plan" (id=2243920) — 226 electrical fixtures + 50 keynotes
**Subject:** Lindenhurst, NY — view "Reed P testing View" — my placements

## Self-score: **60 / 100**

The big structure (tiered placement, wall-snap, cardio power-bar rule, BCS specialty layouts, keynote # table, view-scoped queries) is largely correct. Many specifics — load names, family types, room numbering, and missing fixture categories — are wrong relative to Charlotte's canonical work.

---

## What Charlotte got right that I got wrong in Lindenhurst

### 1. Load names (I used the wrong strings)

| Charlotte | What I used in Lindenhurst | Note |
|---|---|---|
| `CHECK-IN COUNTER` | `CHECK-IN RECEPT` | Different load name |
| `IT SERVER RACK` | `IT RACK` | Different load name |
| `CHECK-IN TABLETS` | `USB RECEPT` (for the check-in USB cluster) | The reception USB cluster has a specific load, not generic USB |
| `BATHROOM SINK` (4) + `MENS RESTRM SINKS` (6) | `BATHROOM SINKS` (lumped 10) | Trough sinks split by men's vs women's restroom |
| `RED WAVE - 103B` | `LAY-DOWN TANNING - 103G` | Different equipment entirely — RedWave is its own load |
| `SAUNA - 103H` | (not placed) | Missing equipment type |
| `BACKWRAP RCPTS` | (not placed) | Missing |
| `RECEPTION GEN. RCPTS` | (not placed) | Missing |
| `I.T. RECEPT` | (not placed) | Missing — distinct from `IT SERVER RACK` |
| `TMAX RECEPT` (1 in Charlotte) | (4 in Lindenhurst) | Charlotte uses 1 TMAX recep; Lindenhurst's 5-in-column convention may be wrong |
| `HYBRID TANNER - 103D` and `- 103F` (per room) | `HYBRID TANNER - 103B` (same load for both hybrids) | Each tanning room has a UNIQUE room number suffix |
| `STAND-UP TANNER - 103C` and `- 103E` (per room) | `STAND-UP TANNER - 103C` (only one) | Same — each room gets its own suffix |
| `TANNING BED - 103K`, `- 103L` | (not placed by room number) | Charlotte uses room-number-suffixed naming for each tanning bed |

### 2. Family types — TV variants

Charlotte uses **`EF-U_Receptacle_CED : Duplex Wall - TV`** for every TV recep:
- TV - BLACK CARD SUITE
- TV - DIGITAL MEDIA
- TV - MEN/WOMEN LOCKER ROOM
- TV & RADIANCE MONITOR - CHECK-IN 102

I used `Duplex Wall` (plain) for most TVs and `Duplex Wall - GFCI` for locker TVs. **All TV receps should be `Duplex Wall - TV`** regardless of which TV they serve.

### 3. TV TRUSS layout — 1 recep per truss, not 2

Charlotte has 20 TV TRUSS receps for what looks like ~20 trusses → **1 Quad Wall - TV per truss**.

I placed 2 receps per truss (1 Duplex Wall - TV + 1 Quad Wall - TV) for a total of 30. **The TV TRUSS load is a single Quad Wall - TV per truss**, not a duplex+quad pair.

### 4. Cardio keynote #13 — multiple, not one TYP

Charlotte has **4 × keynote #13** for cardio — one per cardio group (treadmill row 1, treadmill row 2, stairmaster row, bike row), not a single TYP at the cluster center like I placed.

### 5. TV TRUSS keynote #14 — single, not two

Charlotte has **1 × keynote #14**. I placed 2 (west group + east group). Should be a single TYP across all trusses.

### 6. Hydromassage count — 4 chairs

Charlotte has 4 HYDROMASSAGE - 103A + 4 HYDROMASSAGE RECEPT - 103A (4 chairs total). Lindenhurst has 3 chairs in CAD, so 3 is correct for Lindenhurst — but I should verify this on every project, not assume the playbook count.

### 7. Keynote numbers I missed entirely

Charlotte uses: **#10 (×1)**, **#17 (×1)**, **#18 (×1)**, **#20 (×3)**, **#22 (×1)** — I don't have any of these in my keynote table. They're for fixtures I either didn't place (BACKWRAP, SAUNA, RECEPTION GEN, I.T. RECEPT) or for cases I haven't encountered.

### 8. RECEPT - BCS count

Charlotte has 11 RECEPT - BCS receps (Duplex Wall + Duplex Wall - GFCI mix). I only kept 5 from the takeover; should have placed more along the BCS counter and audited the type (some need GFCI, some don't).

### 9. Total fixture count

Charlotte E101: 226 fixtures. My Lindenhurst: ~170. I'm low — partly because Lindenhurst is smaller, but also because I missed fixture categories.

---

## What I got right (worth keeping in heuristics)

1. **Tier ordering** (A → B → D → E → F → C → G with C moved to end)
2. **Wall-snap with slide-along-wall fallback** for blocked positions
3. **View-scoped FilteredElementCollector** when deleting view-specific elements (after the 3265-keynote disaster)
4. **Sloan trough 5 GFCIs + 2 hand-dryer JBs** layout — correct count matches Charlotte (BATHROOM SINK 4 + MENS RESTRM SINKS 6 = 10 total = 5 per restroom ✓, HAND DRYER × 4 ✓)
5. **Cardio power bar rule** — only TREAD/STEP/HM5A get receps, on the small square in the power bar
6. **Cardinal-snapped rotation** with TV-angled exception
7. **The "longest in band" wall-pick algorithm** for BCS specialty
8. **A-DEMO walls excluded** from active wall set
9. **Disconnects must sit AGAINST walls, not ON them**
10. **Keynote # canonical table** is mostly correct (where I had numbers)

---

## Questions to add to heuristics

I need answers to lock these in for the next project:

1. **TV TRUSS — confirm 1 Quad Wall - TV per truss?** (Not 2 — drop the Duplex Wall - TV from the placement). Or is the type/count truss-dependent (some trusses get extra USB, etc.)?

2. **All TV receps use `Duplex Wall - TV` family type?** Including locker, BCS, digital media, radiance/check-in? (Currently the playbook says only TV TRUSS uses Duplex Wall - TV.)

3. **Tanning room numbering** — Charlotte uses unique 103-letter per room (D, E, F, K, L). How do I figure out which letter goes to which room? Is it numbered going clockwise from the BCS entry, or alphabetically by equipment type, or just whatever's drawn on the architectural plans?

4. **Why is "RED WAVE" a load name in Charlotte but "LAY-DOWN TANNING" is what the yaml says for `PF Tanning 42-3 TLT PLT`?** Should I rename the yaml profile's load name to "RED WAVE - 103X" instead, since the equipment IS a RedWave? Or are they truly different things and I should look for a separate RedWave CAD block?

5. **SAUNA load** — where does the sauna equipment come from in CAD? Is it a `BCS_Sauna` block or similar? I don't think I saw one in Lindenhurst.

6. **BACKWRAP RCPTS** — what's the trigger? Is it a CAD block (e.g., a counter graphic in the BCS or check-in)? Or just convention near a backwrap location?

7. **CHECK-IN TABLETS** vs **USB RECEPT** — when do I use which name? Charlotte uses CHECK-IN TABLETS for the 6 USB receps at reception. Is USB RECEPT a different fixture entirely, or just my wrong name?

8. **CHECK-IN COUNTER vs CHECK-IN RECEPT** — Charlotte uses CHECK-IN COUNTER. Is that the correct canonical name? My playbook has CHECK-IN RECEPT.

9. **Keynote #10 / #17 / #18 / #20 / #22** — what loads do these belong to? (#20 has 3 instances in Charlotte — what's the load?)

10. **TMAX RECEPT** count — Charlotte has 1 single TMAX recep. The "5 receps with middle skipped if TV" convention may be obsolete or specific to a particular site type. Should the Lindenhurst-style 5-stack-minus-1 still apply, or only on certain projects?

11. **RECEPT - BCS count** — how do I determine how many BCS counter receps? Spacing along the counter every X ft? Charlotte has 11; I have 5 (takeover's). What's the rule?

12. **Cardio keynote #13** — is it one TYP per cardio sub-group (treadmill row 1, treadmill row 2, stair row, bike row = 4 keynotes) or one TYP for the whole cardio cluster? Charlotte shows 4 per group.

13. **TV TRUSS keynote #14** — single one for ALL trusses on the floor, regardless of west/east groups? Charlotte shows 1 total.

Once these are answered, the playbook + heuristics get updated and the next project starts from a much stronger baseline.
