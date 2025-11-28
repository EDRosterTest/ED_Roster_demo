# ED_Roster_demo
Python-based emergency department rostering system using constraint programming (OR-Tools) and AI-assisted code generation. Includes input template and sample outputs.

🏥 AI-Assisted Emergency Department Rostering

Overview

This repository contains the source code and example templates for an AI-assisted rostering system developed for a Hong Kong emergency department. The project demonstrates how clinicians can use ChatGPT and Google OR-Tools to build a fully functional, constraint-based duty roster generator without formal programming training.

The system uses constraint programming (CP) to assign shifts (A, P, N, O) according to coverage, fairness, and rest-day rules, and includes a post-processing step that refines outputs into department-specific duty codes.

⸻

⚙️ Features • Python + Google OR-Tools-based constraint solver • Multi-level constraint hierarchy (fixed, adjustable, soft) • Automatic post-processing • Fairness optimisation using penalty weights • Integration with Excel for easy import/export • Adjustable manpower and rest-rule parameters • AI-assisted code development via ChatGPT prompts

⸻

🧩 File Structure

AI-ED-Rostering/ -main.py # Core roster generator (A, P, N, O backbone) -Roster_input.xlsx # input template for demonstration (with loose constraints set) -requirements.txt # Python dependencies -README.md # Documentation

Output_samples/ -Roster_input.xlsx # input template for demonstration (partially titrated constraints) -Roster_Output1.xlsx. # output sample for demonstration ((A, P, N, O backbone) with even distribution of the fairness metrics -Roster_Output2.xlsx # output sample for demonstration ((A, P, N, O backbone -> post processing)

⸻

🧠 Method Summary 1. Backbone generation: The solver first creates a roster composed solely of A, P, N, and O shifts based on user-defined constraints. 2. Constraint hierarchy: • Fixed hard constraints: safety & policy rules (e.g. post-night rest, one duty/day) • Adjustable hard constraints: manpower per shift, duty requests, weekend off, etc. • Soft constraints: fairness, rest balance, penalty-based optimisation 3. Post-processing: The system translates the backbone roster into specific departmental codes (A2, B, E2, D2, etc.) to improve coverage.

⸻

🧮 Requirements • Python 3.9+ • Google OR-Tools • OpenPyXL • Pandas

To install:

pip install -r requirements.txt

⸻

▶️ Usage 1. Edit example_input.xlsx to include your dummy staff list and desired parameters. 2. Run the solver (typing in terminal):

python main.py

3.	Review the generated output
4.	(Optional) Apply post_processing.py for department-specific duty translation.
⸻

How the solver runs

Files & toggles • Input: • Roster_input.xlsx (working sheet, tab Sheet1) • Outputs: • Roster_Output1.xlsx (solver write-back) • Roster_Output2.xlsx (after post-processing) • Roster_Output3.xlsx (transcribed into departmental template)

• Run by typing: "python main.py" in terminal

• Toggles in Sheet1 • cell D3: "Y" turns on soft-penalty optimization (PA/PAN/PPP, etc.) • cell D4: integer stage toggle • 1 → stop after solver write-back (Roster_Output1.xlsx) • 2 → run post-processing and save Roster_Output2.xlsx • 3 → also transcribe into template (Roster_Output3.xlsx)

Prepare your Excel input • Open Roster_input.xlsx (Sheet1) • Adjust constraints if necessary • Manpower section — daily AM / PM / N coverage numbers and senior mix. • Settings (top rows) — min/max Sundays off, weekend off, Sunday PM limits, etc. • Fixed duty requests — mark any pre-decided shifts in the calendar grid (e.g. A, P, O, AL, ☆, A↗). • Optional: adjust hour targets, PA ratio limits, or pattern caps in the side columns. • Save the file after edits.

⸻

Run the solver • Open Codespace / VS Code terminal • Run: python main.py • Wait for the message: ✅ Written Roster_Output1.xlsx • Review the output • Refine iteratively • Adjust constraints (e.g. loosen min/max Off, relax coverage, tune spacing). • Re-run the script — the solver will regenerate automatically. 💡 Tips: If you get “❌ No feasible solution”, some quotas or coverage may conflict — relax one or two limits and retry.

Write back solver results (Roster_Output1.xlsx)

Post-processing / translation (if toggled on) → Roster_Output2.xlsx

Transcribe into departmental template (if toggled on). Saves Roster_Output3.xlsx. (departmental template was not uploaded due to privacy issue)

⸻

Constraint highlights (what the model guarantees) • Exact daily coverage for AM/PM/N. • Rank-mix balance per day (seniors, CON/AC/HT/BT/E bands). • Fixed requests and OFF-types honored exactly where specified. • Night spacing and ≤6/7 workday rule across the month boundary. • PA/PA-N/PPP caps per staff; 4×PM prohibited; daily PA caps by day • P/A ratio compliance (per-staff and global). • Sunday/Weekend Off min/max, Sunday PM quotas. • Objective (if enabled) minimizes PAN, PA, and 3×PM occurrences.

⸻

🔒 Data Privacy

This repository contains only anonymised demonstration data. No identifiable staff information or real duty records are included. For ethical reasons, clinical or operational use should involve local validation.

⸻

📘 Citation

If you reference or adapt this code, please cite: (to be added before publication)

⸻

📬 Contact

For academic correspondence: (to be added before publication)

⸻

⚠️ Disclaimer

This software is provided for research and educational purposes only. It is not a certified clinical scheduling product. Use at your own discretion.
