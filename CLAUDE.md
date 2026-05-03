# LeadManagement — Project Instructions

## How-To Document Formatting Standard

When creating any how-to, guide, reference, or operational document for this business, always generate a `.docx` file using `python-docx` that exactly matches the Source of Truth visual style. Do not create plain text or markdown versions — always produce a `.docx`.

---

## Color Palette

| Name            | Hex       | Usage                                              |
|-----------------|-----------|----------------------------------------------------|
| Dark green      | `#1a3a2a` | All headings, table label column text              |
| Near-black      | `#1c1c1a` | All body text, table detail column text            |
| Gold            | `#c8a96e` | Heading 1 bottom border line                       |
| Cream line      | `#e2e0d8` | Heading 2 bottom border line                       |
| Table header bg | `#2c1f0e` | Dark brown background on table header rows         |
| White           | `#ffffff`  | Table header row text                              |
| Code block bg   | `#F2F2F2` | Light gray background behind code snippets         |

---

## Typography

| Element    | Font        | Size  | Bold | Color     |
|------------|-------------|-------|------|-----------|
| Heading 1  | Arial       | 18 pt | Yes  | `#1a3a2a` |
| Heading 2  | Arial       | 14 pt | Yes  | `#1a3a2a` |
| Heading 3  | Arial       | 12 pt | Yes  | `#1a3a2a` |
| Body text  | Arial       | 11 pt | No   | `#1c1c1a` |
| Code block | Courier New |  9 pt | No   | `#1c1c1a` |

---

## Section Dividers (Bottom Borders on Headings)

Applied via raw XML — python-docx does not support paragraph borders natively.

- **Heading 1**: gold bottom border — `color=c8a96e`, `sz=12`, `space=4`, `val=single`
- **Heading 2**: light gray bottom border — `color=e2e0d8`, `sz=4`, `space=4`, `val=single`
- **Heading 3**: no border

---

## Spacing (in twips — 1 pt = 20 twips)

| Element   | Space Before | Space After |
|-----------|-------------|-------------|
| Heading 1 | 400 twips   | 200 twips   |
| Heading 2 | 300 twips   | 160 twips   |
| Heading 3 | 220 twips   | 120 twips   |
| Body text | 0 twips     | 120 twips   |
| Bullets   | 0 twips     | 80 twips    |
| Code block| 60 twips    | 60 twips    |
| Tables    | 60 twips inside cells (top/bottom), 115 twips (left/right padding) |

---

## Table Formatting

- **2 columns**: Item (1.75 in) and Details (4.5 in)
- **Header row**: dark brown bg `#2c1f0e`, white bold text
- **Item column** (left): dark green `#1a3a2a`, bold, Arial 11pt
- **Details column** (right): near-black `#1c1c1a`, regular, Arial 11pt
- **Cell padding**: top/bottom = 80 twips, left/right = 115 twips
- **Border color**: `#D0CEC6` (light warm gray), single, sz=4
- Add a small spacer paragraph after every table

---

## Document Structure Pattern

Every how-to doc should follow this structure:

1. **Title block** (centered, 22pt, dark green, bold) — two lines: DOCUMENT NAME + SUBTITLE
2. **Subtitle** (centered, 11pt, italic, near-black) — one-line description
3. **"What's inside this document"** table — Section | Contents
4. **"How to use this document"** paragraph — brief intro
5. **Sections** using Heading 1 (`Part X — Title` or `Section X — Title` format)
6. **Steps** using Heading 2 (`Step 1 — Do This` format)
7. **Sub-topics** using Heading 3
8. **Tables** for any structured config, reference, or comparison content
9. **Bullet lists** for step items, confirmations, checklists
10. **Code blocks** (Courier New, gray bg) for any code, URLs, or variable names

---

## Page Margins

- Top: 1 inch
- Bottom: 1 inch
- Left: 1.1 inches
- Right: 1.1 inches

---

## Naming Convention

Save reformatted or new how-to docs in `/Users/creighbaby/LeadManagement/` with descriptive names. Use em dashes (—) in filenames where the Source of Truth does.

---

## Reusable Python Helper (reference)

The full python-docx helper functions (add_h1, add_h2, add_h3, add_body, add_bullet_item, add_code, add_table, set_cell_margins, add_bottom_border, set_para_spacing, set_cell_shading, set_table_borders) are defined in the reformatted Lead Intake doc generation script and should be reused verbatim for all future docs.
