# C55-Onboarding-Comparison-to-MDI
Compares new project data in C55 with the project data from MDI

This version combines:
-Manual column mappings
-Exact matching
-Normalized matching
-Fuzzy matching
-Ambiguous match detection
-Column reconciliation matrix
-Field-level difference reporting
-Duplicate key detection
-Missing-record reporting
-Excel output with multiple worksheets

What this produces

Summary
-Counts of records, mappings, and differences.

Column Mapping Matrix
-Every MDI column.
-Best-matched C55 column.
-Match type (Manual, Exact, Fuzzy).
-Confidence score.
-Acceptance status.

Review Required
-Ambiguous fuzzy matches that need manual approval.

Rejected Matches
-Low-confidence mappings.

Field Differences
-Every value change found between matched records.

MDI Only / C55 Only
-Missing Shared Cost Numbers.

MDI Duplicates / C55 Duplicates
-Duplicate key audit worksheets.

This version is for a formal MDI-to-C55 migration validation because it creates an auditable mapping matrix and prevents questionable fuzzy matches from being automatically compared.
