---
id: ced-electools-sync-one-line-data
doc_type: tool
title: Sync One-Line Data
summary: Copies current panel and circuit data into compatible CED one-line detail components.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Sync One Line
navigation_path: []
status: production
audience: [electrical]
model_impact: updates-one-line-detail-data
keywords: [one-line, panels, circuits, parameters, synchronization]
aliases: [Sync One Line]
last_verified: "2026-08-24"
---

# Sync One-Line Data

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| CED one-line diagram components need current circuit and panel values. | Updates matching parameters on compatible one-line detail components. | Model circuit/panel data and compatible CED one-line components. |

## What it does

Synchronizes selected circuit and panel information from the Revit model to compatible CED one-line detail items.

## When to use it

Use it after changing circuit names, ratings, wire or conduit data, panels, or after running calculations.

## Before you start

- Confirm circuit and panel data is current; run [Calculate Circuits](calculate-circuits.md) when applicable.
- Confirm the one-line components use the required CED parameters and matching panel/circuit identifiers.

## Steps

1. On **AE pyTools > Electrical**, click **Sync One Line**.
2. Choose the required scope or options when prompted.
3. Run the sync.
4. Review representative one-line components against the corresponding model circuits and panels.

## Results and verification

Matching values are copied into compatible detail-item parameters. Verify load name, breaker, wire, conduit, panel, and other required one-line values visually after synchronization.

## Notes and limitations

- The command supports CED one-line families and parameter conventions; unrelated detail items are not automatically supported.
- Circuit data depends on matching panel and circuit-number parameters. Panel data depends on matching panel-name parameters.

## Related pages

- [Calculate Circuits](calculate-circuits.md)
- [Panel Report](panel-report.md)
- [CED ElecTools](index.md)
