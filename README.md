# SA Grants Helper

A lookup-only helper for Regional Fund committee members. It answers questions about the assigned South Asia proposals (Round 3, FY 2026-27) using only the proposal text, with an exact quote and a link to the section on Meta-Wiki. It does not evaluate, score or recommend.

## What's inside

- skills/grant-proposal-lookup/SKILL.md: the rules the assistant follows
- skills/grant-proposal-lookup/references/index.md: list of proposals with links and the exact revision copied
- skills/grant-proposal-lookup/references/proposals/: one file per proposal, copied word for word from Meta-Wiki on 2026-11-22
- skills/grant-proposal-lookup/references/budgets/: one budget file per proposal (same filename), converted from the linked budget spreadsheet on 2026-11-22; originals in budgets/raw/

Not included: talk pages and other linked documents.

## Install

### Claude Code

In Claude Code, run:

    /plugin install sa-grants-helper --marketplace gopavasanth/sa-grants-helper

(On older Claude Code versions, run `/plugin marketplace add gopavasanth/sa-grants-helper` first, then `/plugin install sa-grants-helper@sa-grants-helper`.) Restart Claude Code if asked, then start a new session.

### ChatGPT desktop app or Codex

1. Download the plugin: `git clone https://github.com/gopavasanth/sa-grants-helper.git ~/Downloads/sa-grants-helper`
   (or on GitHub click Code > Download ZIP and unzip it to ~/Downloads/sa-grants-helper).
2. In ChatGPT (Work mode) or Codex ask:

       @plugin-creator Add the plugin folder at ~/Downloads/sa-grants-helper to my personal marketplace so I can install and test it.

   (In Codex, use $plugin-creator.)
3. Restart the app, open Plugins, install "sa-grants-helper", and start a new chat.

## Try it

- "In the Bangladesh proposal, who is on the team and what are their roles?"
- "What dates does the Magh Mela proposal run, and how much does it request?"
- "Does the DCW proposal mention any partners?"
- "Which of these proposals start in February 2027?"
- "What is the travel cost in the Magh Mela budget?"

Every answer should include a quote and a section link. If one doesn't, click through and check.

## Updating

Applicants can edit their pages. Each proposal file records the revision it was copied from. Re-download if pages change before your review. Budget files record their spreadsheet link and download date; to refresh one, download the sheet again as .xlsx into budgets/raw/ with the same filename and run:

  pip install openpyxl
  python scripts/xlsx2md.py skills/grant-proposal-lookup/references/budgets skills/grant-proposal-lookup/references/proposals
