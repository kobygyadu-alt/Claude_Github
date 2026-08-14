# Executive Slide Redesign — Prompt Template

A shared standard for asking an AI assistant (Claude, Copilot, etc.) to redesign
PowerPoint slides to a boardroom-ready, consulting-firm standard with Proforest
branding. Copy the template, fill in the four fields at the top, attach your
`.pptx`, and send.

Why a template instead of freestyling: the difference between a mediocre AI
redesign and a great one is almost entirely in the constraints. This prompt
locks in the things that go wrong most often — brand colours, tables, scope
creep on your wording, and decorative "AI slop" — so results are consistent
across the team.

---

## The master prompt (copy everything in the block below)

```text
Redesign the attached PowerPoint slide(s) to a polished, executive-level
standard in the style of a top-tier consulting firm (McKinsey, BCG, Deloitte).

CONTEXT — read this first
- Audience: [e.g. ExCo / CFO / Project Board / client steering committee]
- Key message: [the single insight or decision each slide must drive — one line per slide]
- Content scope: [RESTYLE ONLY — keep all my wording exactly as written]
  OR [RESTYLE + REWRITE — you may sharpen titles, condense bullets and tighten
  takeaways, but flag every change that alters meaning]
- Slides in scope: [all / slide numbers]

STORY AND STRUCTURE
- Every slide must pass the 10-second boardroom test: a senior executive
  should grasp the point without explanation.
- Give each slide an action title that states the insight or decision, not
  the topic (e.g. "Q3 costs are 12% over budget, driven by contractor rates",
  not "Q3 Cost Overview"). If content scope is RESTYLE ONLY, propose action
  titles separately rather than changing mine.
- Order content top-left to bottom-right in the order I want it read:
  headline, supporting takeaway, then evidence.
- If a slide is trying to make two points, split it into two slides.

BRANDING — Proforest
- Fonts: Aptos (fallback Calibri / Segoe UI). Strong contrast between
  headings (bold, larger) and body (regular).
- Core palette: Turquoise #318E93 (headings, emphasis, key figures),
  Dark Grey #404040 (body text), White #FFFFFF background,
  Light Tint #EDF5F5 (subtle fills, table header rows, section shading).
- Status / variance colours only where they carry meaning:
  Green #38A836 (positive / on track), Yellow #DABA38 (watch / amber),
  Red #A3352F (negative / at risk), Bright Blue #2961A3 (secondary accent).
- Use at most 3 palette colours per slide. Colour draws attention to what
  matters; everything else stays neutral.
- Keep the Proforest logo where it is in the source; do not stretch,
  recolour, or move it.

TABLES
- Transform tables into visually structured, easy-to-scan formats with clear
  hierarchy: light-tint header row (#EDF5F5, bold dark-grey text), generous
  row padding, thin horizontal rules only (no heavy gridlines or boxed cells).
- Right-align all numbers with consistent decimal precision and thousands
  separators. Left-align text. Make totals bold with a top border.
- Show status with small coloured dots or up/down arrows, not fully filled
  cells.
- If a table is too dense to scan, break it into sections, or pull the 2–4
  numbers that matter into KPI callout tiles above it and simplify the rest.

LAYOUT AND TYPOGRAPHY
- Consistent margins (minimum 0.5"), a clear underlying grid, and aligned
  edges across all elements. Balanced whitespace — no clutter.
- Establish hierarchy through size, weight, and spacing, not colour fills or
  boxes around everything.
- Use simple visual elements (icons, dividers, subtle shading, KPI tiles)
  only where they genuinely guide attention or aid comprehension.

DO NOT
- No decorative accent lines under titles or full-width colour bars.
- No cream/beige backgrounds — white or brand colours only.
- No gradients, drop shadows, clip-art, or stock imagery.
- No decoration for its own sake — every visual element must earn its place.
- Never change my numbers, dates, or names under any scope setting.

OUTPUT
- Deliver an editable native PowerPoint file: real text boxes, shapes, and
  tables — never flattened images of slides.
- Preserve the source slide dimensions and aspect ratio.
- After delivering, list per slide: what you changed and why, any copy you
  rewrote, and any assumptions you made that I should verify.
```

---

## Quick version (single simple slide, no tables)

```text
Redesign the attached slide to an executive, consulting-firm standard using
Proforest branding: Aptos/Calibri fonts, Turquoise #318E93 for headings and
emphasis, Dark Grey #404040 body text on white, max 3 colours. Keep all my
wording exactly as written. Give it a clear hierarchy through size, weight
and whitespace; consistent margins; no decorative bars, beige backgrounds,
or clutter. It must pass a 10-second boardroom read. Audience: [WHO].
Key message: [ONE LINE]. Deliver an editable .pptx, and list what you
changed and any assumptions.
```

---

## Field guide

| Field | What to write | Why it matters |
|---|---|---|
| **Audience** | Who will see it (ExCo, CFO, client board) | Sets formality and how much detail survives |
| **Key message** | The one sentence each slide must land | The AI designs *toward* this — it drives the title, emphasis, and what gets cut or highlighted |
| **Content scope** | `RESTYLE ONLY` or `RESTYLE + REWRITE` | The single biggest source of bad surprises. RESTYLE ONLY guarantees your words are untouched; RESTYLE + REWRITE lets the AI sharpen copy but forces it to flag meaning changes |
| **Slides in scope** | "all" or specific numbers | Stops the AI reworking slides you didn't ask about |

If you skip a field, the assistant will usually ask — filling them in up front
saves a round trip.

## Before you send it to the audience

A 60-second human check on the AI's output:

- [ ] Read each title alone — do the titles tell the story by themselves?
- [ ] Spot-check every number, date, and name against your source
- [ ] Numbers right-aligned, consistent decimals, totals distinct
- [ ] Colour appears only where it means something (status, variance, key figure)
- [ ] Nothing overflows, overlaps, or sits off-grid (check in Slide Show view)
- [ ] Review the AI's listed assumptions and rewrites — accept or revert each

## Tips

- **One slide's key message per line.** For a 10-slide deck, list ten one-line
  messages. Slides without a stated message come back generic.
- **Iterate on one slide first.** For big decks, ask for slide 1 only, correct
  course, then request the rest "in the same style".
- **Claude users:** if you're on the Proforest Claude workspace, the
  `executive-slide-redesign` and `proforest-theme` skills apply this standard
  automatically — just attach the deck and describe what you want. This
  template exists so the same standard travels to any tool and any teammate.
