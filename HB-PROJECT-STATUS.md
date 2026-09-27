# BattleTech Historical Battles — Project Status

## Project Architecture

**BattleTech Historical Battles (HB)** is the umbrella project for historically grounded BattleTech campaign work.

**Battle of Tukayyid — 3052** is the currently established campaign/workstream. The `SANSd20/historical-battle-of-tukayyid` repository is the current durable implementation repository for that campaign. This does not require future HB campaigns to use this repository, and no additional campaign is established here.

**BattleTech BSA Tracker** is a separate project. Its authority, task state, data, repository, and priorities are not part of HB. No current HB dependency on it is established.

## Durable Authority

- `README.md` — Tukayyid project description, release state, rules-path summary, and repository map.
- `scripts/build_campaign_rules.py` — editable source used to build the campaign-rules document.
- `docs/Historical_Battle_of_Tukayyid_Campaign_Rules.docx` — formatted v0.2 campaign-rules document.
- `PLAYTESTING.md` — playtest procedure and durable v0.2 playtest priority and resume guidance.
- `release-notes-v0.2.md` and release `v0.2-playtest` — release history and published artifact.
- `docs/Historical_Battle_of_Tukayyid_Forum_Post_SMF.txt` — distribution/forum copy.
- GitHub issue forms — structured playtest, balance, and rules-clarification intake.

This status file is a concise HB umbrella and active-state record. It does not replace the campaign rules, README, source material, force-builder documentation, RAT data, or a future task ledger.

## Current Work and Resume Points

### Campaign-Rules Playtest

Version 0.2 remains the current playtest release. The durable playtest priority is to complete a **Star-scale Region**, with particular attention to **Flank** and **Supply**. Resume using the procedure and evidence checklist in `PLAYTESTING.md`.

### RAT / Availability Review

The active Tukayyid force-building work includes reviewing units and configurations appearing in official Tukayyid-specific Random Assignment Tables against the current unit-availability source.

**RAT/availability review active — exact durable resume point still requires recovery**

No new RAT finding or unit-date determination is established by this status record.

## Tukayyid RAT / Availability Discrepancy Policy

For Battle of Tukayyid — 3052 force-building work, an official Tukayyid-specific RAT entry is sufficient authority for that unit/configuration to appear in the force builder even when the current unit-availability source gives a later Intro Year or availability date than May 3052.

When this occurs:

- retain the official Tukayyid RAT result;
- do not invalidate or reroll it solely because of the later current availability date;
- record the later date as a source discrepancy;
- do not describe the unit/configuration as a prototype unless a source explicitly establishes prototype status; and
- do not generalize the discrepancy into a universal BattleTech availability rule.

This is a project-specific Tukayyid force-building and source-reconciliation policy. It does not modify official BattleTech rules or publications.

## Force-Builder / RAT Durable Location

The active force-builder, RAT records, and supporting data have a verified durable location in external private project storage and remain part of the active Tukayyid/HB work.

The detailed private-storage locator is intentionally omitted from this public repository. Reliable continuation requiring the precise locator should use a separate private durable record. No such private record is created by this status update.

This repository remains the authority for implemented campaign rules. This status record does not alter or reinterpret the privately stored force-builder or supporting data.

## Rules and Source Relationships

The existing campaign-rules hierarchy remains unchanged:

1. The project campaign document changes, replaces, or adds to published rules within this campaign.
2. Where rules conflict, the project campaign document controls within this campaign.
3. Track-specific rules override general campaign rules where established.
4. Required campaign sources and the supported Classic BattleTech and Alpha Strike paths remain those established by current durable campaign documentation.

Project interpretations and force-builder reconciliation decisions must remain distinguishable from published BattleTech rules and historical facts.

## Known Durable-State Gaps

- The exact RAT/availability review resume point still requires recovery and durable recording.
- The precise force-builder/RAT storage locator is not recorded in this public repository and requires a separate private durable record for reliable cross-session discovery.
- The campaign source/generator contains the Alpha Strike rules path, while the current README and forum-post required-publication tables omit its row. This remains outstanding documentation work; it is not corrected here.
- Earlier conversation material concerning force-builder structures, formations, RAT work, and deferrals remains **FORENSIC / HISTORICAL — not current durable project state** unless separately validated and promoted.

## Future HB Task Ledger

The intended future HB umbrella task ledger is `PROJECT-TASKS.md`, governed by `PROJECT-TASK-LEDGER-STANDARD.md` v1.0 in `SANSd20/Project-Management-Standards`.

Because Tukayyid is currently the only durably established HB campaign, that ledger may initially reside in this repository if a later adoption-readiness check confirms the placement. This does not make the repository the permanent home of every future HB campaign.

`PROJECT-TASKS.md` has not yet been created.
