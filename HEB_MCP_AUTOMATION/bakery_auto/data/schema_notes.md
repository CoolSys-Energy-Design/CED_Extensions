# BAKERY auto-design — schema & environment notes

Generated during discovery on 2026-05-19.

## Documents (both open in one Revit session)
- SOURCE (read from): `RunUpdateProfilesHEBCarrollton95PercentCLOUDMODEL`
  - path: `Autodesk Docs://_CED Revit Development Project/RunUpdateProfilesHEBCarrollton95PercentCLOUDMODEL.rvt`
- TARGET (build into): `CED HEB Test Run_MEPR_R24`
  - path: `Autodesk Docs://_CED Revit Development Project/CED HEB Test Run_MEPR_R24.rvt`
- Links present: `Arch`, `Equip`, `00000_Carrollton_v24_Equip`
- View (both projects): `Power Callout - BAKERY - L1`  (Electrical Power Plan, FloorPlan, scale 96)
- Source BAKERY view id: 8141848  (target view id: 8141848 as well per get_current_view_info — verify per doc)

## CRITICAL: spatial filtering
`FilteredElementCollector(doc, viewId)` returns view-associated elements that are NOT
spatially clipped. Wires/circuits span the whole electrical model. Filter every element
to the view crop-box rectangle.

Source BAKERY crop box (world XY, axis-aligned here):
- min: [-69.23, -3.41], max: [-22.79, 83.67]  (feet)
- Robust test: `local = cropTransform.Inverse.OfPoint(P); cb.Min.X<=local.X<=cb.Max.X and cb.Min.Y<=local.Y<=cb.Max.Y`

## Element anatomy
- Electrical Fixtures: family e.g. `EF-U_Receptacle_CED`; LocationPoint; FacingOrientation,
  HandOrientation; Host often null (face/workplane based); MEPModel.GetElectricalSystems()
  -> ElectricalSystem with params RBS_ELEC_CIRCUIT_PANEL_PARAM, RBS_ELEC_CIRCUIT_NUMBER.
- GA_Keynote Symbol_CED: category "Generic Annotations", 23 in source view. (annotation symbol)
- Wires: `.NumberOfVertices`, `.GetVertex(i)` -> XYZ; WireType.
- Circuits: categories "Electrical Circuits" (979) + "Electrical Spare/Space Circuits" (592)
  — most are whole-model; keep only those whose connected elements are in the crop.

## Coordinate mapping strategy (per user)
For the first run the background is identical so absolute XY transfers. The DURABLE goal:
each device/keynote/wire is captured RELATIVE to the nearest linked equipment element it
serves (which equipment by family/type/mark + offset vector + relative orientation), so it
can be reproduced when equipment sits at different relative locations in the target.

## Safe IronPython helpers
- Name: `DB.Element.Name.GetValue(e)` with getattr fallback (plain `.Name` raises AttributeError on some types).
- No auto transaction; wrap modifications in DB.Transaction.

## Agent coordination decision
One Revit instance + single-threaded API => collection scripts run SERIALLY, not parallel.
Collection scripts live in ../skills and are run via MCP; outputs are JSON in this folder so
replication is fully offline-reproducible. Parallel sub-agents are used only for offline
analysis / skill-doc authoring where there is no Revit contention.
