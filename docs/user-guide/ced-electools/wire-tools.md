---
id: ced-electools-wire-tools
doc_type: tool
title: Wire Tools
summary: Creates, redraws, and tags electrical wires and homeruns in plan views.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Wire Tools
navigation_path: []
status: production
audience: [electrical]
model_impact: creates-and-deletes-wires-and-tags
keywords: [wires, homeruns, circuits, tags, interconnect, wire type]
aliases: [Wire Circuited Elements, Select Home Runs]
last_verified: "2026-08-24"
---

# Wire Tools

Wire Tools creates and tags electrical wiring in project floor plans and reflected ceiling plans. It supports circuit-driven branch wiring, direct device interconnects, individual homeruns, and wiring devices to one selected node.

> [!NOTE]
> Select main-model elements and choose a loaded project **Wire Type** before running a wiring operation. Circuit-based schemes resolve selected electrical systems; direct schemes require usable electrical connectors.

## Select a wiring scheme

| Scheme | Use it when | Result |
|---|---|---|
| **Wire by Circuit** | Devices are already circuit-assigned and each circuit network needs branch wiring. | Generates branch wiring for every eligible circuit represented by the selection, including all devices on those circuits. **Skip Single-Device Circuits** is available for this scheme. |
| **Interconnect** | Devices should display as a connected network, including devices on different circuits. | **Full Circuits** includes all devices on the selected circuits and bridges circuit groups; **Selected Only** connects only the selected devices in a spatial nearest-neighbor sequence. Both use one configurable homerun. |
| **Individual Homeruns** | Each selected device needs a separate homerun. | Creates one homerun from each selected device's primary connector. |
| **Wire to Node** | Devices should terminate at one equipment connection or other electrical node. | Wires every selected device to the chosen node and adds a homerun from that node. |

> [!IMPORTANT]
> **Wire to Node** requires compatible connector types on the node and every selected device. Select the node with **Select Node** before running.

## Make a valid selection

Use **Select Devices** to make a filtered Revit selection, or **Use Current Selection** to validate what is already selected.

- For **Wire by Circuit**, the command resolves selected MEP elements or electrical-system objects to circuits of the chosen **System Type**.
- The other schemes need family instances with usable electrical connectors.
- When the selection includes more than one system type, choose a specific **System Type** or retain **All System Types**. The filter applies to every scheme: it filters circuits for **Wire by Circuit** and **Full Circuits**, and connector candidates for direct schemes.
- **Full Circuits** includes every device on selected circuits of the selected system type.
- The status line identifies selected, valid, and invalid items; invalid ones are excluded from the run.
- Use **Refresh Target View** after changing views. It sets the new target view and clears selected devices, node, and homeruns.

## Set wire and geometry options

- **Wire Type** is required for wiring and homerun tagging.
- **Branch Wiring** uses **Arc** or **Chamfer** geometry for circuit branches, interconnects, and device-to-node wires.
- **Homerun Wiring** uses **Arc** or **Chamfer** geometry for custom homeruns.
- **Direction** uses **Toward Panel** when the panel can be resolved, or **Device Facing** to follow the selected device's facing direction.
- **Homerun Length** is the device-connector-to-open-end distance. The default is `4.0`; values are passed as Revit model distance, normally feet.
- **Bend Offset** of `0` produces a straight route. Positive values increase the bend on the calculated side; negative values mirror it. It can also affect custom interconnect paths.

## Redraw existing wires deliberately

**Redraw Active-View Wires Connected to Selected Devices** is enabled by default. Before rebuilding, it removes active-view wires in the operation's scope:

- **Wire by Circuit** removes wires associated with selected circuits.
- Direct and interconnect schemes remove wires connected to selected devices.
- **Wire to Node** removes selected-device wires and the selected node's existing homerun.

> [!WARNING]
> Redraw can remove wires created by a different scheme. Disable it when existing wires must be preserved. It never removes wire tags, so inspect tags after rebuilding.

## Tag existing homeruns

Homeruns can be tagged without running a wiring scheme:

1. Choose a **Tag Type**, **Existing Tags** behavior, and whether to **Add Leaders**.
2. Click **Select Homeruns** to collect open-ended homeruns in the active view.
3. Click **Tag Homeruns** and review the result.

Only active-view wires with exactly one open connector qualify.

- **Skip already tagged** leaves existing tags in place and excludes those wires.
- **Replace existing tags** replaces safe single-reference tags. Multi-reference tags are retained and reported rather than deleted.
- With **Add Leaders** turned off, the command puts a no-leader tag clear of the open endpoint using a scale-aware offset.

> [!TIP]
> The PDF-era **Select Home Runs** command now lives here as **Select Homeruns** and **Tag Homeruns**.

## Review results and resolve issues

Review created and deleted wires, homeruns, skipped circuits, invalid selections, and failures in the status/result output. Each circuit or device operation is isolated where possible, so one failed item does not necessarily stop the complete run.

| Problem | What to check |
|---|---|
| Tool is unavailable | Use a project floor plan or reflected ceiling plan. |
| No eligible circuits | Confirm that selected items belong to circuits of the chosen **System Type**. |
| No shared connector type | Reduce the selection or choose a **System Type** every device supports. |
| Wire to Node fails | Confirm compatible connector types on the node and devices. |
| No homeruns are found | Confirm the wire is open-ended, in the active view, and not excluded as already tagged. |
| Tags remain after redraw | Redraw intentionally leaves tags; use **Tag Homeruns** to review or replace them. |

## Related pages

- [Find Circuited Elements](find-circuited-elements.md)
- [Calculate Circuits](calculate-circuits.md)
- [CED ElecTools](index.md)
