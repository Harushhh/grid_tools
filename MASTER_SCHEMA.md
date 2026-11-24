# MASTER_SCHEMA - how the upload CSV's columns are meant to be filled

For the agent that will maintain and extend this automation (read it together with `AGENT_CONTEXT.md`). The user gave this schema on 26 Oct 2026 "for the understanding of the agent who will
in future work on this automation": it explains the LOGIC of every column of the portal's rate-config upload. It is an explanation, not a change request - the five scripts in
`D:\harsh\grid_tools` keep the conventions they have (the user confirmed "we can keep it as it is"), and part 4 says where those conventions look different from the wording of the schema, so
you do not mistake them for bugs.

---------------------------------------------------------------------------------------------------------------------------------------

## Part 1 - The schema (the user's rules, in their words)

> This schema defines the strict mapping logic for converting insurer rate grids into the finalized, uploadable CSV format. Any automated script must enforce these rules to remain fault-tolerant.

### 1. Identifiers & global master mappings
- `id`: leave blank (handled manually post-export).
- `new_vehicle_makes`: MUST contain the mapped Make/Model cluster name from the global **Make Model Master**.
- `new_rto_list`: MUST contain the mapped RTO cluster name from the global **RTO Rate Master**. *(Superseded by the user on 30 Oct 2026: the scripts now write only the cluster name as the grid writes it, e.g. `GJ_ABS`, and the user maps and groups the RTOs himself - `AGENT_CONTEXT.md` section 9.)*
- `insurer_vertical`: default to `Not Applicable`.
- `insurance_company`: must exactly match the specific insurer's dropdown validation name.
- `user_id`: hardcoded admin ID for the automated script.

### 2. Product, policy & vehicle specifications
- `product`: broad category (e.g. `Two-Wheeler`, `GCV 4-Wheeler`, `Private Car`).
- `sub_product`: coverage structure (e.g. `1+1`, `1+3`, `SAOD`, `STP`).
- `policy_type`: `Renew` if sub_product is `1+1`, `SAOD` or `STP`; `New` if sub_product is `1+3` or `1+5`.
- `fuel_type`: `Petrol`, `Diesel`, `CNG`, `LPG`, etc.
- `vehicle_age_min` / `vehicle_age_max`: numeric range calculated from the manufacturing year to the current year.
- `make_model_class`: specific vehicle sub-type (e.g. `Car`, `Bike`, `Truck`). **Wildcard:** if the grid gives a flat rate for the entire product without specifying sub-classes, output `NA`.
- `cc_min` / `cc_max`: engine capacity or tonnage range. If unmentioned in the grid, default to the wildcard `0` to `99999`.
- `sc_min` / `sc_max`: seating capacity range. If unmentioned in the grid, default to the wildcard `0` to `999`.
- `veh_use`: leave blank.

### 3. Pricing In (PI) rates & limits
- `pi_type`: strict router; the outputs are mutually exclusive:
  - `On TP`: populate only `pi_tp_rate` (leave OD / Net blank).
  - `On OD`: populate only `pi_od_rate`.
  - `On Net`: populate only `pi_net_rate`.
- `pi_tp_2`, `pi_tp_3`, `pi_tp_4`, `pi_tp_5`: additional TP rates (if applicable).
- `pi_flat_amount`: default to `0`.
- `pi_vli`: hardcode to `0`.
- `tariff_min` / `tariff_max`: OD discount range. The absolute maximum cap for `tariff_max` is `100.01`.

### 4. Pricing Out (PO) rates & broker margins
- `po_type`: MUST always mirror `pi_type` exactly.
- `po_od_rate`, `po_tp_rate`, `po_net_rate`, `po_flat_amount`: calculated by subtracting a configurable margin variable `x` from the corresponding `pi` rate.
  Example: `pi_type` = `On OD` and `pi_od_rate` = `17` -> `po_od_rate` = `17 - x`. `x` must be defined dynamically so it can be updated in the future.

### 5. Dates, flags & edge cases
- `from_date`: the explicit effective date provided by the insurer on the new grid (e.g. `01-07-2026`).
- `to_date`: default to `31/03/2030`. When a new grid arrives, the old grid's `to_date` updates to the new grid's `from_date` minus 1 day.
- `is_ncb` / `is_cpa` / `is_zd` (No Claim Bonus, Personal Accident, Zero Depreciation): if the grid explicitly splits rates for these conditions, create separate rows for `Yes` and `No`;
  if the grid does not mention them, default to `NA`.
- `add_tnc`: free-text field to capture highly specific grid conditions / rules that cannot be mapped into the standard columns.
- `remarks`: the dedicated exception log. Use this column strictly for the fallback mechanism (e.g. logging "UNMAPPED RTO" or "UNMAPPED MAKE"). The script MUST NOT crash on missing
  master mappings; it must log them here and continue processing.

---------------------------------------------------------------------------------------------------------------------------------------

## Part 2 - The ideas behind it (what to carry in your head)

1. **A row is one rate card.** The portal matches a policy against rows; a row says "for THIS insurer, product, sub_product, fuel, make group, RTO group, age band, cc/GVW band, seat band, NCB/CPA/ZD
   situation, discount band and T&C, the insurer pays pi, we pay po". Everything not stated by the grid is a WILDCARD, never an invented restriction: `NA` for the class and the flags, `0`-`99999` for cc,
   `0`-`999` for seats, `0`-`99` for age, `0`-`100.01` for the discount.
2. **Names come from the masters, never from the grid.** The grid says "DELHI - NCR"; the portal only knows the group `LIBERTY_PC_TW_APR26_DELHI_-_NCR`. A name the master lacks is logged, not
   replaced by a lookalike (`AGENT_CONTEXT.md` section 3).
3. **`pi_type` decides which one rate column is used** (`On TP` -> `pi_tp_rate`, `On OD` -> `pi_od_rate`, `On Net` -> `pi_net_rate`), and `po_type` repeats it. That is how the portal knows the
   premium base the percentage is applied to. Never fill two rate columns of one row.
4. **`po = pi - x`**: `x` is the broker's margin, a setting that may change (`PO_CUT`, 7 points today in every script); the scripts never let a payout go below 0.
5. **Dates chain**: a new grid starts on its own `from_date`; the old grid's rows should end the day before it; `to_date` is the far end (`31/03/2030`) while a grid is current.
6. **Split rows, do not blur**: if a grid pays differently with/without NCB, CPA or Zero Dep, that is two rows (`Yes` / `No`), not an average; if the grid is silent the flag is `NA`.
7. **Fault tolerance**: a mapping the master does not have, a note that cannot be turned into logic, a payout with no portal sub_product - none of these may crash the run or be guessed. They are logged in
   `remarks` (or, for a grid problem, put to the person as a question) and the run continues, so the user can fix the exception by hand.

---------------------------------------------------------------------------------------------------------------------------------------

## Part 3 - Column by column: the schema, and what the scripts write today

"Same" = the scripts do what the schema says. Where the convention differs, it is the one the portal ACCEPTED in the user's own uploads, which the scripts were built to reproduce (regression
baselines): do not change it unasked.

| Column | Schema | The scripts today |
|---|---|---|
| `id` | blank | `1` on every row (Liberty, RSA GCV, RSA Private Car, United), `1..N` (Magma) - as in the user's reference uploads |
| `new_vehicle_makes`, `new_rto_list` | master cluster names | Same. `check_names` (shared block, all five scripts) logs a name the master lacks and carries on |
| `insurer_vertical` | `Not Applicable` | Same |
| `insurance_company` | exact dropdown name | Same: `LIBERTY GENERAL INSURANCE LTD`, `MAGMA HDI GENERAL INSURANCE CO LTD`, `ROYAL SUNDARAM GENERAL INSURANCE COMPANY LIMITED`, `UNITED INDIA INSURANCE CO LTD` |
| `user_id` | hardcoded admin id | `USER_ID` setting: `2966` (Liberty, Magma, RSA); empty in United (as in its July upload) |
| `product` | e.g. `Two-Wheeler`, `GCV 4-Wheeler`, `Private Car` | the portal's spellings from the uploads: `TW`, `GCV 4W`, `GCV 3W`, `PCV 4W`, `PCV 3W`, `PCV 2W`, `MISCD`, `Private Car` |
| `sub_product`, `policy_type` | `1+1` / `1+3` / `SAOD` / `STP`; Renew vs New | Same, plus `1+5` (New, two wheelers). `3+3` and long-term liability have no portal sub_product: not written, flagged |
| `fuel_type` | Petrol, Diesel, CNG, LPG, etc. | those + `Electric`; "non-diesel = petrol" grids give one row per fuel with the same rate |
| `vehicle_age_min/max` | range from manufacturing year to current year | the age BAND the grid gives (`0-1.01`, `1.01-5.01`, ...), `0-99` when the grid has none. Turning a policy's manufacturing year into an age is the portal's mapping engine's job: it compares that age with these bands |
| `make_model_class` | sub-type; `NA` when flat | as in the uploads: `Car`, `BIKE`, `SCOOTER`, `Taxi`, `School Bus`, `Other Bus`, `Goods Carrying Rickshaw`, `Passenger Carrying Rickshaw`, `PCV 2W`, `MISCD-Tractor`, `MISCD-Others`; `NA` for GCV 4W and (United) TW |
| `cc_min/max`, `sc_min/max` | wildcards `0`-`99999`, `0`-`999` | Same; bands use `previous limit + 0.01` (`1000.01`-`1500.01`), open end `99999` / `999` |
| `veh_use` | blank | Same |
| `pi_type` and the rate columns | one populated, the others blank | one rate column non-zero and `pi_type` names it; the unused ones hold `0` (as in the uploads) |
| `pi_tp_2..5`, `pi_flat_amount`, `pi_vli` | extra TP rates; `0`; `0` | `0` in all of them |
| `tariff_min/max` | cap `100.01` | `0`-`100.01`; RSA GCV uses its discount bands (`0-81.01` and `81.01-86.01`, EAST `0-85.01` and `85.01-90.01`) |
| `po_type`, `po_*_rate` | mirrors `pi_type`; `pi - x` | Same with `x` = `PO_CUT` = 7.0 (a setting in each script), floored at 0 (zero stays zero) - the floor is the user's standing rule of 23 Oct |
| `from_date` | the insurer's effective date | typed each run (DD/MM/YYYY) and compared with the "w.e.f." date in the document; a mismatch is flagged |
| `to_date` | `31/03/2030`; old grid ends the day before the new one | typed each run: `31/03/2030` for the Liberty / Magma / RSA / United-October files, month ends for RSA Private Car September and United July / August. Nothing end-dates the previous grid's rows yet |
| `is_ncb`, `is_cpa`, `is_zd` | `Yes` / `No` rows when split, else `NA` | Same logic; spelling follows each insurer's upload: Magma `Yes` / `No`, RSA `YES` / `NO`, Liberty and United `NA` |
| `add_tnc` | free text for conditions that cannot be mapped into standard columns | part of the row key: fixed T&C texts the portal matches on (`Ambulances`, `EX:Ambulances`, `New,<text>`, `Old,<text>`, the school-bus text, the Magma CPA text) |
| `remarks` | exception log (UNMAPPED RTO / UNMAPPED MAKE); never crash | `check_names` writes `new_rto_list 'X' is not in the RTO master` / `new_vehicle_makes ...` (with the master's spelling as a hint) and carries on. The same column also carries `Manual Review Required: [note]` for notes, text rules and missing annexures a script cannot apply (user's standing rule, 23 Oct 2026) |

---------------------------------------------------------------------------------------------------------------------------------------

## Part 4 - Worked examples (real rows from the October 2026 United upload, empty columns left out)

A two-wheeler package row, payout floored at 0 (pi 5 - 7 < 0), with a manual-review flag from a note about a policy type that has no portal sub_product:
```
product TW | sub_product 1+1 | policy_type Renew | fuel_type Petrol | make_model_class NA | new_vehicle_makes TW_ALL | new_rto_list AS,KA,KL,MP,PY,TN
age 0-99 | cc 0-150.01 | sc 0-999 | pi_net_rate 5 | pi_type On Net | tariff 0-100.01 | is_ncb NA is_cpa NA is_zd NA | from_date 26/10/2026 | to_date 31/03/2030
po_type On Net | po_net_rate 0 | remarks "Manual Review Required: [Two Wheeler Long Term Package (5+5): the portal's July upload has no sub_product for it ...]"
```
A Private Car standalone third-party row (`STP` -> `On TP`: only the TP column carries the rate; po = 17.5 - 7):
```
product Private Car | sub_product STP | policy_type Renew | fuel_type Petrol | make_model_class Car | new_vehicle_makes Private_Car_ALL | new_rto_list ALLINDIA
cc 0-1000.01 | pi_tp_rate 17.5 | pi_type On TP | po_type On TP | po_tp_rate 10.5 | from_date 26/10/2026 | to_date 31/03/2030
```
A Private Car own-damage-only row (`SAOD` -> `On OD`; the payout 25 - 7 = 18):
```
product Private Car | sub_product SAOD | policy_type Renew | fuel_type Diesel | cc 0-99999 | pi_od_rate 25 | pi_type On OD | po_type On OD | po_od_rate 18
```
A goods vehicle row whose payout cites an annexure that is not in the document - written for the whole states, flagged (`AGENT_CONTEXT.md` rule 11):
```
product GCV 4W | sub_product 1+1 | new_vehicle_makes GCV_4W_ALL | new_rto_list HR,MP,TN,UP | cc (GVW kg) 2000.01-3500.01 | pi_net_rate 15 | pi_type On Net | po_net_rate 8
remarks "Manual Review Required: ['(Except RTOs listed in Sub Annexure 2)' for 2000-3500 kg: that annexure is not in this document, so it is not applied ...]"
```
A split-by-NCB row (Magma, own-damage private car; the NCB case is its own row, the no-NCB case another with `is_ncb No`; add_tnc is part of the key):
```
product Private Car | sub_product 1+1 | fuel_type Diesel | new_rto_list MAGMA_APR26_AP | pi_od_rate 19.25 | pi_type On OD | is_ncb Yes | is_cpa NA
add_tnc "CPA Premium will not be considered for outgo if less than Rs 450" | po_od_rate 12.25 | id 16417
```

---------------------------------------------------------------------------------------------------------------------------------------

## Part 5 - Rules for the agent

1. **Do not "correct" the scripts toward the wording of the schema.** Where Part 3 shows a different convention (a `1` in `id`, a `0` in an unused rate column, `TW` for the product, `Yes`/`YES`, the dates typed per
   run, `Manual Review Required` notes in `remarks`), the user reviewed it and decided to keep it (26 Oct 2026). The schema explains the intent behind the columns; the scripts' conventions are what the portal accepted.
2. **For a NEW insurer or product** follow the schema; where it contradicts the user's own answer-key CSV for that insurer, the CSV's convention is the one the portal accepted - use it, say so in the report,
   and ask once only if neither is clear.
3. **Fault tolerance is not optional**: a missing master mapping is logged and processing continues; a missing master FILE, an unreadable cell or an ambiguous grid is a question (or a stop under `--no-prompt`),
   never a silent guess and never a crash with a traceback.
4. If the user later asks to change a convention, change the script(s), re-run `check_regressions.py`, re-bless only with the user's OK, update Part 3 here, and record the decision
   (`learn.py note --section decision`).
5. See also `AGENT_CONTEXT.md`: section 2 (the output contract as the scripts write it), section 3 (master lists), section 4 (the remarks protocol), section 9 (decisions already made).
