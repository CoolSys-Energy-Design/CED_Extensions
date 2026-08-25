# Wire Tools

Wire Tools creates and tags electrical wires in a floor plan or reflected ceiling plan. It supports circuit-based wiring, direct interconnection, individual homeruns, and wiring to a selected node.

The tool works in project documents only. Select main-model elements and a loaded project **Wire Type**. Connector-based schemes require usable electrical connectors; circuit-based schemes resolve the selected system type.

## Wiring schemes

| Scheme | Use it when | What it creates |
|---|---|---|
| **Wire by Circuit** | Devices are already assigned to circuits and should be wired using their circuit networks. | Revit-generated branch wiring for every eligible circuit represented by the selection. All devices on those circuits are included. **Skip Single-Device Circuits** is available here. |
| **Interconnect** | Devices should appear as one connected network, including across circuits. | **Full Circuits** includes all devices on the selected circuits and adds bridges between circuit groups. **Selected Only** connects only the selection in a spatial nearest-neighbor sequence. One configurable homerun is created. |
| **Individual Homeruns** | Each selected device needs its own homerun. | One homerun per selected device, using its primary connector. |
| **Wire to Node** | Several devices should terminate at one electrical node or equipment connection. | A direct wire from each selected device to the node, plus a homerun from the node. Devices and node must have compatible connector types. |

When the selection contains more than one system type, choose the type from **System Type**. It defaults to **All System Types**, so mixed selections remain valid until you choose a specific target. The selector applies to every scheme: it filters circuits for **Wire by Circuit** and **Full Circuits**, and filters connectors for the direct wiring schemes. **Full Circuits** includes all devices on selected circuits of the chosen type.

## Selection and run behavior

- **Select Devices** opens a filtered Revit selection. **Use Current Selection** validates the current Revit selection instead.
- For **Wire by Circuit**, selected MEP elements or electrical-system objects are resolved to circuits of the selected system type after selection.
- For the other schemes, selected elements must be family instances with usable electrical connectors.
- For **Wire to Node**, select one node with **Select Node** before running.
- The status line shows selected, valid, and invalid elements. Invalid items are excluded from the run.
- If the active view changes, use **Refresh Target View**. This changes the target view and clears device, node, and homerun selections.

## Wire and geometry options

- **Wire Type** — The Revit wire type applied to created wires. Required for wiring and homerun tagging.
- **Branch Wiring** — **Arc** or **Chamfer** for circuit branches, interconnects, and device-to-node wires.
- **Homerun Wiring** — **Arc** or **Chamfer** for custom homeruns.
- **Direction** — **Toward Panel** aims toward the circuit panel when it can be resolved. **Device Facing** follows the selected device's facing direction.
- **Homerun Length** — Distance from the device connector to the open endpoint. The default is `4.0`; values are passed directly as Revit model distance, normally feet.
- **Bend Offset** — `0` creates a straight path. Positive values increase the bend on the calculated side; negative values mirror it. This can also affect custom interconnect paths.

### Redraw existing wires

**Redraw Active-View Wires Connected to Selected Devices** is on by default. It deletes active-view wires in the selected scope before rebuilding:

- Wire by Circuit: wires associated with the selected circuits.
- Direct/interconnect schemes: wires connected to selected devices.
- Wire to Node: selected-device wires and the node's existing homerun.

It can delete wires created by another scheme. Turn it off when existing wires must be preserved. Wire tags are not deleted by redraw, so review tags after rebuilding.

## Homerun tagging

Wire Tools can tag existing open-ended homeruns without first running a wiring scheme:

1. Choose a **Tag Type**, **Existing Tags** behavior, and whether to **Add Leaders**.
2. Click **Select Homeruns** to find open-ended homeruns in the active view.
3. Click **Tag Homeruns**.

Only homeruns in the active view are collected, and a wire must have exactly one open connector.

- **Skip already tagged** leaves existing tags in place and excludes those wires from selection.
- **Replace existing tags** replaces safe, single-reference tags. Multi-reference tags are left in place and reported rather than deleted.
- With **Add Leaders** off, a no-leader tag is placed clear of the open endpoint using a scale-aware offset.

## Results and troubleshooting

Each circuit/device operation is isolated where possible, so one failure does not necessarily stop the rest of the run. Review created wires, deleted wires, homeruns, skipped circuits, invalid selections, and failures in the status/result output.

| Problem | Check |
|---|---|
| Tool is disabled | Use a floor plan or reflected ceiling plan in a project document. |
| No eligible circuits | Confirm the selection belongs to circuits of the selected **System Type**. |
| No common connector type | Reduce the selection or choose a **System Type** supported by every device. |
| Wire to Node fails | Confirm the node and devices have compatible connector types. |
| No homeruns found | Confirm the wires are open-ended homeruns in the active view and are not excluded as already tagged. |
| Tags remain after redraw | Redraw does not remove tags; use **Tag Homeruns** to review or replace them. |
