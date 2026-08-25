---
id: ced-mechtools-transition
doc_type: tool
title: Transition
summary: Creates a transition fitting between two compatible open connector ends.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > Utilities > Transition
navigation_path: [Utilities]
status: beta
audience: [mechanical]
model_impact: creates-fittings-and-moves-elements
keywords: [transition, fittings, connectors, piping, ductwork]
aliases: []
last_verified: "2026-08-24"
---

# Transition

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Two compatible open MEP ends require a transition fitting. | Creates a transition and can move the first picked element to accommodate it. | Two open, compatible connectors in the same domain. |

## What it does

Prompts for two open connector ends, creates an appropriate transition between them, and adjusts the first selected element as required by the transition length.

## When to use it

Use it to connect compatible MEP ends of different sizes or geometry with a transition rather than manually placing and adjusting the fitting.

## Before you start

- Confirm both ends have unused connectors in the same domain.
- Pick near the exact connectors to use.

## Steps

1. On **AE pyTools > Mechanical > Utilities**, click **Transition**.
2. Pick the first element near the connector that may move.
3. Pick the second, static element near its target connector.
4. Review the created transition and connected geometry.

## Results and verification

A transition is created between compatible connectors. Verify fitting type, length, element movement, connector continuity, and system assignment.

## Notes and limitations

> [!WARNING]
> This beta command changes geometry and can move the first picked element. Do not use it on a broad selection or without reviewing the result.

## Related pages

- [ConnectTo](connect-to.md)
- [Make Parallel](make-parallel.md)
- [CED MechTools](index.md)
