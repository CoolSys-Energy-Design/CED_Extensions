---
id: ced-mechtools-set-piping-system
doc_type: tool
title: Set Piping System
summary: Assigns a selected piping or duct system type and can create a system for eligible unsystemed segments.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > Set Piping System
navigation_path: []
status: production
audience: [mechanical]
model_impact: modifies-and-creates-mep-systems
keywords: [piping, ductwork, system type, connectors, MEP systems]
aliases: [SetPipeSystem]
last_verified: "2026-08-24"
---

# Set Piping System

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Selected piping or duct segments need a different system type or need to be systemed. | Can divide, retype, reconnect, or create MEP systems. | Select a coherent piping or duct network and choose the target system type. |

## What it does

Applies a selected piping or duct system type to the selected network. For eligible unsystemed segments, it can create a system instead of only changing an existing system type.

## When to use it

Use it to correct system assignment after modeling changes or to establish systems for a completed connected network.

## Before you start

- Select only the piping or duct segments that belong in the target workflow.
- Confirm connected boundary elements and the target system type.
- Save or synchronize before changing a shared model.

## Steps

1. Select the piping or duct segments to update.
2. On **AE pyTools > Mechanical**, click **Set Piping System**.
3. Choose the target piping or duct system type.
4. Review the output and inspect the resulting network.

## Results and verification

The command applies the target type to existing systems or creates systems for eligible unsystemed networks. Verify system browser membership, connector continuity, system type, and affected schedules.

## Notes and limitations

> [!WARNING]
> This command can temporarily disconnect and reconnect network boundaries and can divide systems. Use a focused selection and verify the entire affected network.

## Related pages

- [ConnectTo](connect-to.md)
- [Disconnect](disconnect.md)
- [CED MechTools](index.md)
