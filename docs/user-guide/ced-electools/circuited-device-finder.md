---
id: ced-electools-circuited-device-finder
doc_type: tool
title: Circuited Device Finder
summary: Finds devices connected to selected circuits and selects them in the model.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Tools > Circuited Device Finder
navigation_path: [Circuit Tools]
status: production
audience: [electrical]
model_impact: none
keywords: [electrical, circuits, devices, finder, selection]
aliases: [Circuit Element Finder]
last_verified: "2026-08-24"
---

# Circuited Device Finder

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| You need to locate the devices served by one or more circuits. | Changes the current selection and opens relevant model context. | Selected circuits, or a circuit chosen from the picker. |

## What it does

Collects the devices on selected circuits and selects them in the model. If no circuits are selected, it presents a circuit picker.

## When to use it

Use it to inspect circuit membership before editing, troubleshooting, or coordinating device locations.

## Before you start

- Select the target circuits for a focused result, or be ready to choose them in the picker.

## Steps

1. Select circuits if desired.
2. Run **Circuited Device Finder**.
3. Choose circuits when prompted.
4. Review the selected devices in Revit.

## Results and verification

Confirm that the selected devices match the intended circuit(s) before using the selection for another workflow.

## Notes and limitations

- The command does not edit circuit membership or device properties.
- It requires the CED electrical library to be available.

## Related pages

- [Find Circuited Elements](find-circuited-elements.md)
- [Circuit Manager](circuit-manager.md)
