---
id: ced-mechtools-connect-to
doc_type: tool
title: ConnectTo
summary: Moves one connector-based element to a selected compatible connector and connects the two.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > ConnectTo
navigation_path: []
status: production
audience: [mechanical]
model_impact: moves-and-connects-elements
keywords: [connectors, piping, ductwork, connection, move]
aliases: []
last_verified: "2026-08-24"
---

# ConnectTo

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| An element with an unused connector should be aligned and connected to another open connector. | Moves and may rotate the first picked element, then connects its chosen connector. | Two distinct connector-based elements with compatible unused connectors. |

## What it does

Prompts you to pick an element to move and then a target element. It chooses the unused connector nearest each picked point, aligns the moved element, and connects the two connectors.

## When to use it

Use it to complete a direct connection without manually moving and rotating the first element into place.

## Before you start

- Identify two distinct elements with unused, compatible connectors.
- Pick close to the intended connector on each element.

## Steps

1. On **AE pyTools > Mechanical**, click **ConnectTo**.
2. Pick the element to move near its intended connector.
3. Pick the target element near its intended connector.
4. Review the moved element and the new connection.

## Results and verification

The first element is rotated or moved as needed and its selected connector is connected to the target connector. Verify location, orientation, and system continuity.

## Notes and limitations

- The two connectors must be in the same domain and be unused.
- Picking the same element twice or an element without a usable connector stops the operation.

## Related pages

- [Transition](transition.md)
- [Disconnect](disconnect.md)
- [CED MechTools](index.md)
