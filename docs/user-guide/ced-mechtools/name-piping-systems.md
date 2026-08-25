---
id: ced-mechtools-name-piping-systems
doc_type: tool
title: Name Piping Systems
summary: Places refrigeration piping text labels from rack start pipes and downstream branch rules.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > Ref Ops > Name Piping Systems
navigation_path: [Refrigeration]
status: beta
audience: [refrigeration]
model_impact: creates-and-updates-text-notes
keywords: [refrigeration, piping, system IDs, text notes, labels]
aliases: []
last_verified: "2026-08-24"
---

# Name Piping Systems

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Refrigeration piping needs consistent System ID labels. | Creates or updates text-note labels on refrigeration piping. | A model with correctly identified rack start pipes and connected refrigeration branches. |

## What it does

Builds piping labels from rack start pipes and downstream branch rules, then places text notes to identify refrigeration systems.

## When to use it

Use it after the refrigeration network is modeled and system identity is ready to be documented.

## Before you start

- Confirm rack start pipes and branch connectivity are correct.
- Open the target annotation view and review existing system labels.

## Steps

1. On **AE pyTools > Mechanical > Ref Ops**, click **Name Piping Systems**.
2. Complete the prompts for the target workflow.
3. Review generated labels throughout the affected piping layout.

## Results and verification

Text notes are placed or updated based on the identified piping relationships. Verify System IDs, label locations, and branch assignments before publishing.

## Notes and limitations

> [!WARNING]
> This beta tool relies on model connectivity and refrigeration conventions. Incorrect rack starts or branch relationships can produce incorrect labels.

## Related pages

- [System Tagger](system-tagger.md)
- [Print Pipe Data](print-pipe-data.md)
- [CED MechTools](index.md)
