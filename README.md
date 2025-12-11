I’ve been building an automated EVMS pipeline that uses both the Cobra exports and the OpenPlan/Penske activity file. Before I lock anything in, I want to make sure I’m interpreting the Cobra files the same way you do today.

From some EDA, I’m seeing three main issues:
	1.	Different Cobra formats
	•	Files like Cobra-Abrams STS.xlsx, Cobra-Abrams STS 2022.xlsx, and Cobra-XM30.xlsx have the “classic” structure with columns such as SUB_TEAM, COST-SET, DATE, HOURS.
	•	Others, e.g. John G Weekly CAP Oly 12.07.2025.xlsx and some Stryker files, are very wide tables with many “Results”/date columns, Control_Acct, Currency, etc., and sometimes no SUB_TEAM.
	2.	Program mapping to Penske
	•	In Penske, PROGRAM has values like ABRAMS STS, XM30, M-SHORAD ILS, MLIDS_C-UAS, etc.
	•	When I try to match Cobra SUB_TEAM or control-account fields directly to Penske SubTeam/control accounts, I get no exact matches, which suggests there’s a mapping or naming standard I don’t have.
	3.	Inconsistent sub-team lists
	•	XM30 has a clean set of sub-teams (e.g., AUT, LS, LTH, MFG, PAT, SE, SUR, SW, VET).
	•	Abrams STS/other programs have dozens of numeric-looking sub-teams (e.g., 2025-1, 2045AA-2, 2073-1, etc.).
	•	For the dashboards, it would be ideal if we could either:
	•	map these back to a standard sub-team list similar to XM30, or
	•	confirm that each program is expected to have its own unique sub-team structure.

To move forward without guessing, could you help me with a few specific points?
	1.	For each Cobra export (e.g., the ones listed above), what is the correct PROGRAM in Penske it should tie to? Is it always one-to-one (file → program), or do any files span multiple programs?
	2.	What is the official join key between Cobra and Penske when you calculate EVMS today?
	•	Is it SUB_TEAM, Control_Acct, WBS, something else, or a separate cross-walk?
	3.	For the non-standard formats (John G Weekly CAP Oly and the key Stryker file(s)):
	•	Which columns represent the CTD Cobra values used for CPI/SPI, vs 4-week, YTD, etc.?
	•	Are these files intended to feed the same EVMS dashboards as Abrams/XM30, or are they for a different purpose?
	4.	Is there a standard list of sub-teams by program (or program family) that we should be using, and a mapping from the more detailed Cobra sub-teams (e.g., 2045AA-2) to those standard buckets?
	5.	Longer term, is there a preferred Cobra export layout (columns/filters) we could request for all programs so the pipeline can use a single, agreed-upon schema?

Once I have your guidance on the program mapping, join keys, and sub-team standards, I can update the pipeline so it mirrors the current manual process instead of making assumptions.

Thanks a lot for helping me get this right.

Best,
Megan
