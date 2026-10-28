# OTEC Dr. Dang schedule template project

Clinic Manager: Charoo (OTEC / OLEC). Goal: redesign Dr. Dang's OSCAR schedule template into a family of templates.
Pipeline: diagnostic test -> OD (TG, GT or DP) 20 min later -> Dr. Dang 20 min after that.

## nd-template-builder.html
Single-file app (vanilla JS, no build step). Tabs: Trends, variants (Standard/Summer/Light/Consult push/Two-OD), Free build, Compare.
- Variants: dropdown + drag to assign slot codes. Laser (L) never goes on OD columns.
- Free build (v2): every appointment is its own block (lanes ND/TG/GT/DP, 5-min snap, own length). Placing an OD block auto-adds a Dr. Dang block +20 min (setting); afterwards blocks move independently. Saved templates in localStorage (key nd-template-builder-v3) with JSON export/import.
- Published artifact: https://claude.ai/artifact/DLkhiYBGrdvrxCbiFswRE9
- Slot codes: r refractive consult, c consult other, a follow-up, po post-op, u urgent, o open, L laser, l LDD, P pre-consult.

## analysis/
Python scripts (pandas, openpyxl, xlrd) that built the ND/MI workbooks. Raw OSCAR exports and pickles are NOT included (patient names/ages = PHI). Refractive rule: CAT/cataract/RLE counts only if it is the PRIMARY (earliest) reason; otherwise "Consult Other".

## Open questions
- Should LDD (l) also stay off OD columns?
- Confirm TG sheet provider (Tarn Gill) and real away dates for Dr. Dang / Dr. Iizuka.
- ICL/RCS/LAL/PRK currently counted as "other", not refractive.
