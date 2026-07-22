# Patch scripts

Each script applies one change to `../USD_SOFR_Curve_Bloomberg.xlsx` and was run once.
They are kept as the build record — an `.xlsx` cannot be diffed, so these are how you see
*what* changed and *why*. Read the docstring at the top of each; that is where the reasoning
lives.

They are **not** idempotent and are not meant to be re-run against the current workbook.
To rebuild from scratch, `git checkout` the pristine file and replay them in commit order.

Paths inside are absolute. That was fine when they lived outside the repo; if you ever replay
them, point `WB` at this repo's copy first.

Notable ones, if you are looking for a specific fix:

| script | what it fixed |
|---|---|
| `patch_curve_longend.py` | circular reference that stopped the curve past 10Y |
| `patch_s490_real_data.py` | real S490 quotes + business-day roll on pillar dates |
| `patch_first_coupon_tau.py` | first coupon accrual ran from spot, not settle |
| `patch_dfspot_node.py` | DF(spot) hardcoded as 1 on the curve grids, 2.08bp on a 1Y swap |
| `patch_swap_grid_hole.py` | a note merged over K39:L40 wiped the 18Y/19Y nodes |
| `patch_observed_on.py` | removed an o/n stub that had been fitted to the answer |
| `patch_d1_d3.py` | settlement date, par spread (3.3), upfront (3.2) |
| `patch_testing_fix.py` | empty strings from the CDS pull taking out the hazard chain |
