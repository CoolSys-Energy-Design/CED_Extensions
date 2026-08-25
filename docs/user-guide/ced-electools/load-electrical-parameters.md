---
id: ced-electools-load-electrical-parameters
doc_type: tool
title: Load Electrical Parameters
summary: Loads the shared parameters required by CED circuit-calculation workflows.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Tools > Load Electrical Parameters
navigation_path: [Circuit Tools]
status: production
audience: [electrical]
model_impact: adds-shared-parameters
keywords: [shared parameters, calculations, circuits, project setup]
aliases: []
last_verified: "2026-08-24"
---

# Load Electrical Parameters

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| A project lacks the shared parameters required for CED electrical calculation tools. | Adds shared parameters to the project. | Confirm the project has not already been configured from the CED template. |

## What it does

Loads the shared parameters used by **Calculate Circuits** and related CED electrical workflows.

## When to use it

Use it once during electrical project setup when the project was not created from a CED-configured template.

## Before you start

- Coordinate with the project owner before adding shared parameters.
- Confirm that required parameter names do not conflict with another project standard.

## Steps

1. On **AE pyTools > Electrical > Circuit Tools**, click **Load Electrical Parameters**.
2. Complete the parameter-loading workflow.
3. Confirm the required parameters are available before running calculations.

## Results and verification

The required shared parameters are added to the project. Verify parameter availability on representative electrical elements and circuits.

## Notes and limitations

> [!IMPORTANT]
> This changes project data. Run it intentionally and coordinate with the model-management team.

## Related pages

- [Calculate Circuits](calculate-circuits.md)
- [Circuit Tools](circuit-tools.md)
