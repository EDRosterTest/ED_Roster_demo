🏥 AI-Assisted Emergency Department Rostering

Python-based constraint programming system using Google OR-Tools and AI-assisted code generation

⸻

📌 Overview

This repository contains the source code and example templates for an AI-assisted rostering system developed for a Hong Kong emergency department. The project demonstrates how clinicians can use ChatGPT and Google OR-Tools to build a fully functional constraint-based duty roster generator without formal programming training.

The system uses constraint programming (CP) to assign core shifts—A (AM), P (PM), N (Night), O (Off)—according to coverage, rest-day, and fairness rules.
A post-processing module then converts these basic duties into department-specific shift subtypes for real-world deployment.

⸻

⚙️ Features
	•	Python + Google OR-Tools constraint solver
	•	Multi-level constraint hierarchy (fixed, adjustable, soft)
	•	Automatic post-processing into department duty codes
	•	Fairness optimisation via penalty weights
	•	Excel-based input/output for easy use
	•	Adjustable manpower, seniority mix, and rest-rules
	•	AI-assisted code creation using ChatGPT prompts
	•	Modular architecture for further refinement

⸻

🚀 How to Run the Rostering Program

You can run this program in two ways:


Option 1 — Run in GitHub Codespaces

(Recommended for non-technical users; requires a free GitHub account)
	1.	Log in to your GitHub account.
	2.	Open this repository:
https://github.com/EDRosterTest/ED_Roster_demo
	3.	Click Use this template → Open in Codespaces
or Code → Create Codespace on main
	4.	A cloud-based VS Code session will open with all files pre-loaded.
	5.	Install dependencies (first time only):

pip install -r requirements.txt


	6.	Run the solver by clicking Run ▶, or via terminal:

python solve.py




Option 2 — Run Locally (no GitHub login required)
	1.	Visit the repo (no login required):
https://github.com/EDRosterTest/ED_Roster_demo
	2.	Click Code → Download ZIP
	3.	Unzip the folder
	4.	Ensure Python 3.9+ is installed
(VS Code + Python extension recommended)
	5.	Install dependencies:

pip install -r requirements.txt


	6.	Run the solver:

python solve.py



⸻

🧩 File Structure

This repository contains all essential components for generating a roster.



🔧 Solver - solve.py
	•	Core Python script that generates the roster
	•	Modular coding structure for future extension
	•	Produces the output files when executed



📥 Input Template - Roster_input.xlsx
	•	Main Excel template used by the solver
	•	Sample version represents a 28-doctor November 2025 roster with pre-filled duty requests
	•	Users may adjust quotas, constraints, duty requests, manpower tables, etc.



📤 Output Files (Generated after running the solver)

Roster_Output1.xlsx
	•	Backbone roster with A / P / N / O assignments
	•	Reflects satisfaction of all hard constraints
	•	Includes staff statistics (Sun Off, Weekend Off, Sunday PM, P/A ratio, hour balance)
	•	Includes day statistics (AM/PM/N counts, seniority mix, PA counts)

Roster_Output2.xlsx (output if run mode =2)
	•	Post-processed roster
	•	Converts A/P/N/O into department-specific duty subtypes:
		•	Morning (AM) duties: A (08–16), B (07–15), K (07:30–15:30), A2 (08–17), D2 (09–18)
		•	Evening (PM) duties: P (16–24), E2 (15–24), S2 (15–23)
		•	Night/Others: N (00–08), Z2 (non-clinical), T / ½t (Training)
		•	Special duties (as suffix)
 			•	* (shift IC), ♥ (resus), %¥ (clinic/lab)
			•	^, ⓦ, ω (EM ward related)
			•	O® (reserved off)
      •	Example: A2♥ means 08-17 duty hour with resus duty; E2* means 15-24 hour as shift IC  
		•	Pattern conversions (e.g., P→A becomes S2→A2 or E2→D2)

Roster_Output3.xlsx (Not included for privacy)
	•	Department-format roster rewritten into the official template

📁 Sample Files (Sample/ folder)
	•	Roster_input.xlsx — Demonstration input
	•	Roster_Output1.xlsx — Sample backbone roster
	•	Roster_Output2.xlsx — Sample post-processed roster

📦 Supporting Files
	•	requirements.txt — Python package dependencies
	•	README.md — Documentation

⸻

📘 Input File, Output Files, and Encoded Rules

Below summarises how the input file works and how the solver interprets rules.



📥 1. Input File (Roster_input.xlsx)

The input file contains five main components:

1. Staff Information Table

Defines individual staff-level constraints:
	•	Name, Rank (CON, AC, HT1/HT2, BT, Elective)
	•	Night quotas: N*, N, N3
	•	Night spacing
	•	Sunday Off, Weekend Off, Sunday PM
	•	P/A ratio limits
  •	Target hour balance
	•	Hour range
	•	Limits on PA, PAN, PPP patterns

2. Calendar Grid (Days × Staff)

Users may pre-fill:
	•	A, P, N, O
	•	AL, ☆
	•	noA, noP, noN (prohibitions)
	•	↗ to indicate a staff-requested shift

The solver interprets these as hard constraints.

3. Global Settings

Optional department-wide rules:
	•	Min/Max Sunday Off
	•	Min/Max Weekend Off
	•	Min/Max Sunday PM
	•	Global PA ratio
	•	Global night-spacing requirement

These act as adjustable-hard constraints.

4. Manpower Requirements (Manpower Block)

Daily coverage rules:
	•	Required AM / PM / N headcount
	•	Min/Max seniors
	•	Min/Max CON / AC / HT / BT / E per shift

Defines safe staffing and seniority distribution.


5. Run Modes

Optimisation toggle (cell D3)
	•	“N” — no penalties (faster; feasibility first)
	•	“Y” — apply penalties for unfavourable patterns (searches best roster within 300s)

Module toggle (cell D4)
	•	1 → Solver only (Output1)
	•	2 → Solver + Post-processing (Output1 + Output2)
	•	3 → Full pipeline (Output1 + Output2 + Output3)

📤 2. Output Files (Summary)
	•	Output1: Backbone roster (A/P/N/O)
	•	Output2: Department shift subtypes
	•	Output3: Departmental template (not included)


🧠 3. Key Rules Encoded (Constraint Logic)

A. Fixed Hard Constraints (non-negotiable)
	•	One duty per day
	•	≤6 workdays in any 7-day window
	•	Mandatory A–N–O sequence for night duties
	•	No P→P across Sat–Sun
	•	Required senior mix
	•	At least one specialist in every A/P shift
	•	Honour all pre-filled duties

B. Adjustable Hard Constraints
	•	Staff duty requests (modifiable after discussion)
	•	Daily staffing coverage for A, P, N
	•	Rank-mix minimum/maximum
	•	Night frequency and spacing
	•	Weekend/Sunday Off allocation
	•	Hour-balance range
	•	P/A ratio
	•	Caps for PA, PAN, PPP patterns

C. Soft Constraints

Used when optimisation toggle = “Y”:
	•	Penalties for PA, PAN, PPP
	•	Encourages fairness while preserving feasibility

⸻

💡 Tips for Running the Solver Effectively

Generating a feasible roster is an iterative process. The following workflow is recommended:

1. Start Simple

Begin with:
	•	Minimal fixed requests
	•	Loose constraints (wide min/max ranges)
	•	Fewer restrictions on weekend off, Sunday PM, PA ratio, pattern caps, etc.

Once the backbone roster is feasible:
	•	Check coverage counts
	•	Review seniority distribution
	•	Inspect day-by-day AM/PM/N balance
	•	Verify staff hour balance and P/A ratios

2. Tighten Constraints Gradually

Add or strengthen constraints one group at a time, such as:
	•	Narrowing senior min/max per shift
	•	Tightening PA or PAN caps
	•	Increasing night spacing
	•	Adjusting weekend/Sunday Off distributions
	•	Applying more duty requests

After each adjustment:
	•	Re-run the solver
	•	Ensure feasibility is preserved

This progressive tightening ensures stable convergence without overwhelming the model.

3. Tune Fairness or Penalties Last

Once feasibility is stable:
	•	Turn on optimisation (cell D3 = Y) for penalty weights for PA, PAN, PPP
	•	Apply penalty-based seniority balancing if desired

Penalty functions shape the quality of the roster but may significantly increase runtime.
Use only after the core constraints are functioning well.

4. Handling Infeasibility

If the solver reports no solution:
	1.	Identify the likely bottleneck
	•	Night quotas?
	•	Senior mix limits?
	•	Weekend Off caps?
	•	Too many fixed duty requests?
	2.	Loosen the constraints that are most restrictive
	3.	Re-run until feasibility returns, then continue fine-tuning.

6. Final Optimisation

Once feasibility and general fairness are acceptable:
	•	Run a final optimisation cycle
	•	Review Output2 for correct department subtypes
	•	Use Output3 (if enabled) for operational-format export

⸻

🔒 Data Privacy

Only anonymised demonstration data are included.
No real staff information or clinical data are stored in this repository.

⸻

📘 Citation

Chi-kit Sin, Shu-wing Kung. Implementation and Development Experience of an AI-Assisted Rostering System in a Hong Kong Emergency Department. Hong Kong Journal of Emergency Medicine.
DOI: 10.1002/hkj2.70061
⸻

📬 Contact

Dr SIN, CHI KIT
Department of Accident and Emergency
Tseung Kwan O Hospital
Email: johnsin1113@gmail.com

⸻

⚠️ Disclaimer

This software is intended for research and educational use only.
It is not a certified clinical scheduling product.
Use at your own discretion.
