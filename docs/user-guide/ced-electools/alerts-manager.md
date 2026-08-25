---
id: ced-electools-alerts-manager
doc_type: tool
title: Alerts Manager
summary: Reviews and manages saved electrical circuit alerts in a modeless window.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Alerts Manager
navigation_path: []
status: production
audience: [electrical]
model_impact: Can change saved alert state and recalculate an editable circuit.
keywords: [electrical, alerts, circuits, modeless, recalculate]
aliases: [Circuit Alerts]
last_verified: "2026-08-24"
---

# Alerts Manager

## At a glance

| Use this when                                                                      | Model impact                                                                      | Required context                         |
|------------------------------------------------------------------------------------|-----------------------------------------------------------------------------------|------------------------------------------|
| You need to review saved circuit alerts without opening each circuit individually. | Alert actions change saved alert state; recalculation can change circuit results. | An open project with circuit-alert data. |

## What it does

Opens a modeless manager for inspecting active and hidden circuit alerts, selecting related model elements, and recalculating an editable circuit.

## When to use it

Use it during electrical QA to triage saved alerts and navigate quickly to affected circuits or model elements.

## Before you start

- Save outstanding model changes before recalculating circuits.
- Refresh the manager when the model has changed since it opened.

## Steps

1. Run **Alerts Manager**.
2. Select an alert or circuit to inspect its details.
3. Use the available action to select related elements, change alert state, or recalculate when appropriate.
4. Refresh and confirm the resulting status.

## Results and verification

Verify the alert's updated state and, after recalculation, confirm the circuit values in the model or Circuit Manager.

## Notes and limitations

- The window is modeless and can remain open while you navigate Revit.
- Operations that fail are surfaced in the window; refresh after resolving model conditions.

## Related pages

- [Circuit Manager](circuit-manager.md)
- [Electrical System Check](electrical-system-check.md)
