# Miner Town Follow-up Kit — Traditional Flowcharts

These are **Mermaid** flowcharts (real ovals / boxes / decision diamonds).

### How to see them as a picture
- **Easiest:** open <https://mermaid.live>, delete the sample, and paste one of
  the ```mermaid``` code blocks below — the diagram renders instantly.
- **VS Code:** install the extension *"Markdown Preview Mermaid Support"*, open
  this file, and press the Preview button (top-right).
- **GitHub:** if this repo is on GitHub, it renders automatically.

Shapes used: `([ rounded ])` = start/end · `[ box ]` = action ·
`[/ slanted /]` = file read/write · `{ diamond }` = decision.

---

## Diagram 1 — The full cycle (refresh → build → review)

```mermaid
flowchart TD
    A([Start: time to follow up]) --> RPT[response_report.py<br/>scan Gmail Sent and Inbox]
    RPT --> O1[/Write responded.csv, no_response.csv,<br/>has_book.txt, needs_help.txt/]
    O1 --> BLD[build_control_sheet.py<br/>read data and preserve prior choices]
    BLD --> ROW{For each<br/>recipient}

    ROW --> NEV{Never Follow-up<br/>= yes?}
    NEV -->|Yes| NOQ[Follow-up = No<br/>permanent hard stop]
    NEV -->|No| BLK{Has book, opted out,<br/>or suppressed?}
    BLK -->|Yes| NOQ
    BLK -->|No| IDLE{Idle 7 or<br/>more days?}
    IDLE -->|No| KEEP[Keep prior choice<br/>usually No]
    IDLE -->|Yes| AQ[Auto-queue:<br/>Follow-up = Yes]
    AQ --> TMPL{Already sent them<br/>a template?}
    TMPL -->|No| T1[Template = gentle-nudge]
    TMPL -->|Yes| T2[Template = BLANK<br/>you choose next]

    NOQ --> SHEET[/Write followups.xlsx/]
    KEEP --> SHEET
    T1 --> SHEET
    T2 --> SHEET
    SHEET --> REVIEW[You review and mark:<br/>Yes/No, template, custom, flags]
    REVIEW --> GO([Proceed to Execute])
```

---

## Diagram 2 — Execute: the per-row guardrail gate

```mermaid
flowchart TD
    S([Run execute_followups --send]) --> PRE[PRE-REFRESH:<br/>report + build_control_sheet]
    PRE --> LOAD[Load sheet, suppression, has_book,<br/>cooldown times, sent history]
    LOAD --> EACH{Next row with<br/>Follow-up = Yes?}
    EACH -->|none left| POST[POST-REFRESH:<br/>report + build_control_sheet]
    POST --> DONE([Done])

    EACH -->|yes| G1{Never Follow-up<br/>= yes?}
    G1 -->|Yes| SK1[SKIP: never follow-up]
    G1 -->|No| G2{In suppression.txt?}
    G2 -->|Yes| SK2[SKIP: suppressed]
    G2 -->|No| G3{In has_book.txt?}
    G3 -->|Yes| SK3[SKIP: has the book]
    G3 -->|No| G4{Under 7 days since<br/>last follow-up?}
    G4 -->|Yes| HOLD[HOLD: cooldown<br/>stays queued]
    G4 -->|No| BODY{Custom set?<br/>else Template?}
    BODY -->|neither| ERR[ERROR: pick<br/>template or custom]
    BODY -->|custom| RC{Identical custom<br/>already sent?}
    BODY -->|template| RT{Template sent before<br/>and Allow Repeat = no?}
    RC -->|Yes| BL1[BLOCK: edit the message]
    RC -->|No| SEND
    RT -->|Yes| BL2[BLOCK: choose<br/>another template]
    RT -->|No| SEND

    SEND[Build threaded reply<br/>and SMTP send] --> LOG[/Log send + hash;<br/>Status = SENT;<br/>Follow-up to No; Allow Repeat to no/]
    LOG --> WAIT[Wait 8 to 22 seconds]

    SK1 --> EACH
    SK2 --> EACH
    SK3 --> EACH
    HOLD --> EACH
    ERR --> EACH
    BL1 --> EACH
    BL2 --> EACH
    WAIT --> EACH
```

---

## Diagram 3 — The repeating life of one recipient

```mermaid
flowchart TD
    E([You email someone the original]) --> W{7+ days pass<br/>with no activity?}
    W -->|No| W
    W -->|Yes| Q[Auto-queued for a nudge]
    Q --> CH{They do something?}
    CH -->|Reply meaningfully| OUT([Out: they engaged])
    CH -->|Click 'got it' / say so| OUT2([Out: has the book])
    CH -->|Say no / unsubscribe| OUT3([Out: suppressed])
    CH -->|You mark Never Follow-up| OUT4([Out: never])
    CH -->|Nothing| SENDN[Send next nudge:<br/>1st gentle-nudge, 2nd still-interested]
    SENDN --> SI{Was that the<br/>still-interested nudge?}
    SI -->|Yes| OUT5([Out: Never Follow-up set — stop chasing])
    SI -->|No| W
```
