---
id: ced-mechtools-disconnect
doc_type: tool
title: Disconnect
summary: Disconnects selected pipe or duct elements from unselected connected network elements.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > Disconnect
navigation_path: []
status: production
audience: [mechanical]
model_impact: disconnects-and-divides-mep-systems
keywords: [disconnect, piping, ductwork, connectors, systems]
aliases: [DisConnect]
last_verified: "2026-08-24"
---

# Disconnect

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| A selected portion of a pipe or duct network must be separated from the rest. | Disconnects network-boundary connectors, can divide systems, and removes orphaned selected endpoints from systems. | Select the elements to isolate. |

## What it does

Breaks pipe or duct connections that cross from the selected elements to unselected elements, then updates affected system topology where possible.

## When to use it

Use it before relocating, deleting, or reassigning a branch that must no longer remain connected to its current network.

## Before you start

- Select the full portion of the network to isolate.
- Review downstream systems and coordinate with other users of the network.

## Steps

1. Select the pipe, duct, or connector-based elements to isolate.
2. On **AE pyTools > Mechanical**, click **Disconnect**.
3. Review the output and inspect open connectors and resulting systems.

## Results and verification

Boundary connections are removed. A multiple network can be divided, and fully disconnected selected endpoints can be removed from systems. Verify connector states and system membership.

## Notes and limitations

> [!WARNING]
> This changes topology and system membership. It does not merely remove the current selection.

- The command supports piping and ducting connector domains.

## Related pages

- [ConnectTo](connect-to.md)
- [Set Piping System](set-piping-system.md)
- [CED MechTools](index.md)
