---
id: ae-pytools-set-system-calculation-levels
doc_type: tool
title: Set System Calculation Levels
summary: Sets mechanical and piping system-type calculation levels to None.
extension: AE pyTools
ribbon_path: AE pyTools > Miscellaneous > MiscTools > Set System Calculation Levels
navigation_path: [Miscellaneous]
status: production
audience: [mechanical]
model_impact: modifies-system-types
keywords: [mechanical, piping, system calculations, performance]
aliases: [Turn Off System Calculations]
last_verified: "2026-08-24"
---

# Set System Calculation Levels

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Mechanical or piping system calculations need to be disabled project-wide. | Sets the Calculation Level of mechanical and piping system types to **None**. | Use only with the project team's approval. |

## What it does

Changes the Calculation Level setting on the project's mechanical and piping system types to **None**.

## When to use it

Use it only when the project intentionally disables system calculations, such as a performance-recovery workflow approved by the model owner.

## Before you start

- Confirm that disabling calculations is appropriate for the project.
- Coordinate with mechanical and piping users who depend on calculated system data.

## Steps

1. On **AE pyTools > Miscellaneous > MiscTools**, click **Set System Calculation Levels**.
2. Allow the command to update the system types.
3. Review the Calculation Level of representative system types.

## Results and verification

Mechanical and piping system types are set to **None**. Verify required schedules, calculations, and connected workflows still behave as expected.

## Notes and limitations

> [!WARNING]
> This changes project-wide system-type settings. It can affect downstream analysis and calculations.

## Related pages

- [Miscellaneous Tools](miscellaneous.md)
