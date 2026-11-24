# Grid tools - Liberty and Magma upload files

Two scripts that turn the monthly insurer grids into the rate-config upload CSV:

| Script | Input | Output (in `output\`) |
|---|---|---|
| `liberty_grid_to_config.py` | the Liberty branch grid (.xlsx) | `LIBERTY_UPLOAD_<yyyy-mm>.csv` |
| `magma_grid_to_config.py` | the Magma comprehensive grid (.xlsx) **and** the Magma SATP grid (.xlsb/.xlsx) | `MAGMA_UPLOAD_<yyyy-mm>.csv` |
| `rsa_gcv4w_comp_to_config.py` | the Royal Sundaram (RSA) **GCV 4W comprehensive** grid (.xlsx, one sheet per region) | `RSA_GCV4W_COMP_UPLOAD_<yyyy-mm>.csv` |
| `rsa_pvtcar_to_config.py` | the Royal Sundaram (RSA) **Private Car** grid (.xlsx: four region sheets, the Top End Variant List, the STP sheet) | `RSA_PVTCAR_UPLOAD_<yyyy-mm>.csv` |
| `united_commission_to_config.py` | the United India **"Revised commission structure" PDF** (needs the `pdfplumber` library; `setup.bat` installs it) | `UNITED_UPLOAD_<yyyy-mm>.csv` |
| `shriram_broker_grid_to_config.py` | the Shriram General **broker grid** (.xlsx: the rate sheet plus the Declined-Excluded / PAN India / Conditions sheets) | `SHRIRAM_UPLOAD_<yyyy-mm>.csv` |
| `shriram_2w_to_config.py` | the Shriram General **two-wheeler grid** (.xlsx: New Business - 2W, ROLLOVER 2W (1+1), TW SAOD) and, optionally, the broker grid (its PAN India sheet lists the declined RTOs) | `SHRIRAM_TW_UPLOAD_<yyyy-mm>.csv` |
| `godigit_agency_to_config.py` | the Go Digit General **agency grid** (the password-protected workbook `Delhi_PAN_India_Agency.xls`: commercial vehicles, two-wheelers, Private Car; reading it needs the `msoffcrypto-tool` package, listed in `requirements.txt`) | `GODIGIT_UPLOAD_<yyyy-mm>.csv` |
| `sbi_grid_to_config.py` | the SBI General **broker grid** (Agency, C & D: one wide rate sheet of RTO clusters with a block of columns per product, the High-End enabler, the RTO cluster mapping and the declined make / model lists) | `SBI_UPLOAD_<yyyy-mm>.csv` |
| `icici_grid_to_config.py` | the ICICI Lombard grids - any of the four .xlsx workbooks: the **commercial-vehicle** grid (COMP and AOTP matrices), the **Private Car** grid, the **Two-Wheeler** grid and the **new Two-Wheeler (1+5)** grid (one sheet per make); each is recognised by what its sheets say | `ICICI_UPLOAD_<yyyy-mm>.csv` |
| `indusind_grid_to_config.py` | the IndusInd General (Reliance General) **TC / PO grid for brokers** (.xlsx, the November and the October 2026 layouts: the commercial-vehicle sheet; the Private Car, TW Scooter and Short Term Taxi sheets are listed, not converted yet) - `--from-date` and `--to-date` on every run | `INDUSIND_UPLOAD_<yyyy-mm>.csv` |

RSA is done product by product: each product has its own script (GCV 4W comprehensive, Private Car; the others follow).

### United (PDF): what it does

The circular is a set of bordered tables (Two Wheelers, Passenger Carrying, Goods Carrying, Miscellaneous, Private Car) with an "Existing" and a "w.e.f. <date>" column pair;
the script reads the cell borders, so a class or band label that spans several rows, or one states cell over several payouts, is joined to the right payout by position. It writes the
NEW column: every payout (commission %, states) becomes rows for each fuel of its band - `1+1` On Net, `STP` On TP (two wheelers also SAOD and 1+5), PO = PI - 7. "All States" is
`ALLINDIA`; a list of states becomes state codes; "Other than above States" is every code not named by the other payouts of the band. "(Except RTOs listed in Sub Annexure 3)" and the
preferred RTOs of Sub Annexure 3 use the portal's RTO groups (`ANNEX3_GROUPS` and `RTO_PREFIX` near the top of the script). Goods weight bands are `cc_min` / `cc_max`, bus passenger
bands `sc_min` / `sc_max`.

**Private Car** is built from the table's segments, which together cover every non-electric vehicle exactly once: diesel up to 1500 cc; above 2500 cc except Tata / Maruti / Mahindra / Toyota /
Hyundai / Honda / Kia; and "other" (those 7 makes above 2500 cc, every other non-diesel up to 2500 cc, diesel from 1500 to 2500 cc), for New (1+3), Package (1+1), SAOD and SATP (STP); Electric and SAOD
are "all vehicles". Sub Annexure 4 (40% in the RTOs of 12 cities) is applied to Package (1+1) only, like your July upload (the circular does not say which policy types it covers; you decided on 26 Oct 2026 to keep 1+1 only); the other
policy types are written for ALLINDIA with no remark about it (`CAR_ANNEX4_POLICIES` near the top of the script extends it). A payout that is the same in the Existing and the new column keeps the Existing date
(01/04/2026). Long Term Package (3+3) has no sub_product in the upload, so it is not written (warned, flagged).

**Not converted** (listed as warnings): Sub Annexure 2 (extra incentive by channel). Not written because the July upload has no place for them: TW Long Term 5+5, Motor Trade, Standalone CPA,
Private Car 3+3 (their product's rows carry them in `remarks`). Checked against the July upload: 367 keys match (the 5 Sub Annexure 4 rows differ only in date: 01/07/2026 in the PDF, 01-04-2026 in
the upload); it differs from the PDF for School Bus, PCV 3W (Madhya Pradesh 25%), GCV 3500-7500 kg (Pondicherry), the STP rows of the Sub Annexure 3 groups and in Private Car - the script follows
the PDF and the comparison lists the differences. `user_id` is left empty, as in the July upload. `grid_doctor.py` cannot read PDFs yet.

**August circular (effective 01-08-2026).** Same script, new layout: "Revised Pay-out" columns, a Private Car table by cc x states with a revised commission only (no make groups), Sub Annexure 2 = the
31 Rajasthan RTOs for GCV 2000-3500 kg (56.5%), Sub Annexure 3 = the 40% Private Car RTOs. The portal's RTO groups of the annexures are chosen by the circular's date (`circular_groups`,
`CIRCULAR_GROUPS` near the top of the script): when a new month's circular arrives, add its groups there. **Page 1 of every circular is a scanned cover letter that the script cannot read - it warns; read it
yourself** (it holds rules such as "applicable for all policy types" and the 40% on Net Premium). Run it like July: `--grid <pdf> --from-date 01/08/2026 --to-date 31/08/2026`; the output was written as
`output\UNITED_UPLOAD_2026-08.csv` (511 rows, 163 flagged in `remarks`).

**If the PDF is a scan** (October 2026: every page is a picture) the script stops with "this PDF is a scan" - nothing is wrong with your file. The tables have to be read off the page images into a
transcription `.json` (see `regression/inputs/united_oct26.json` and `read_transcription` in the script) that carries the SHA-256 of the scan; from then on you run the PDF as usual and the script finds its transcription by itself (or pass the `.json` as the grid). An annexure the document does not contain (the October file has no Sub Annexure pages) is NOT applied - nothing is carried over from an earlier month: a payout "except RTOs listed in Sub Annexure 2" is
written for the whole state(s) and flagged in `remarks`, and no 40% preferred-city rows are written.

### Shriram (broker grid): what it does

One sheet of rate rows (state | product | discount % | payout % | policy type | age | remarks). Each row becomes the rows for its product (Private Car, PCV 4W taxi / bus, PCV 3W, GCV 3W,
GCV 4W by weight band, Misc-D): `P` = 1+1, `L` = STP, `P & SA-OD` = 1+1 and SAOD, `Bundle (1+3)` = 1+3. The payout is the payin (`25`, `15 OD`, `16 Net`, `15 OD + 2.5 TP`); the payout to the
agent is 7 points less. The discount % becomes the tariff limit (a higher discount slab starts where the lower one ends). The remarks are read for the cases the script knows (Except New Mahindra, Bolero models, NCB
cases, nil-dep payouts, Dzire / Ambasador taxis, cc ranges left out, tankers, zones, RTO lists ...); a row that prices one make on its own (`Only EECO`, `Bolero Model only`) takes that make out of the other rows of its state;
a note that only says what changed since last month (`PO increased by 2%`) needs no action; every other remark - including a number nobody used - is copied into the `remarks` column as `Manual Review Required: [text]` - never guessed.
If you type the RTO names yourself in a column headed `RTO`, a cell that holds an Excel error (`#REF!`) counts as empty and the master lookup is used for that row (the run says how many).

The RTOs of a row are the state (or Mumbai / ROM / Delhi & NCR ...) **minus the RTOs the sheets "Declined-Excluded" and "PAN India" decline for that product**. The script looks in the master RTO list for
the group that holds exactly the remaining RTOs and writes its name; if the master has no such group the row is written for the whole state and the remark names the group to create
("No RTO group in the master for MP without MP 06, 07, 30"); the same list, with the RTO codes each missing group must hold, is written next to the upload as the sheet `RTO groups to create` of the review workbook. Make names are written as the master spells them. A row that is the same (key and payin) as a row of the earlier upload given with `--ref` keeps that row's from_date. "For Hazardous category PO will be 5% less" is written into the `remarks` column of every row that carries it (the payout stays as listed, as in your working sheet). Electric Private Cars / taxis are not written (the Conditions sheet declines them): run with
`--write-ev` to write them as the earlier uploads did. Two grid rows that give one key two different payouts are both written and both flagged. The short-term policy sheet is not converted.

**STP and duplicate rates.** On STP rows `is_ncb` and `is_zd` are always NA (STP has no own-damage cover). If two grid rows give the same entry different payins (for example the "NCB cases" and "Without NCB" rows of one P & L row, which meet on a single STP entry), the **lowest** payin is written and the entry is listed in the sheet `Lowest rate chosen` of the review workbook.

**Addendum.** Shriram sends an addendum workbook some months (same columns as the rate sheet, rows marked "w.e.f 1st Sep"). Give it with `--addendum <file>` (in Colab: upload it together with the grid). Its rows **override** grid rows: an addendum row replaces the grid rows of the same state whose weight band it covers (a band split in two addendum rows still counts), and only for the policy types it names - an addendum "L" row replaces just the STP part of a "P & L" grid row. A band that overlaps only in part is not trimmed: both rows are written and both say so in `remarks`. Rows with nothing to replace ("New Addition") are added. The run prints what was replaced; without `--addendum` it reminds you to check whether Shriram sent one.

### Shriram (two-wheeler grid): what it does

Two sheets of rate rows. **New Business - 2W** (state | makes | body | discount % | average net payout): each row becomes 1+5 / New rows (bike up to 180 cc, scooty, EV once for bikes and once for scooters), On Net, discount % -> tariff limit.
**ROLLOVER 2W (1+1)** (state | body | ALL EXCL. TVS / TVS | age 6 to 15 | discount | payout | P / L): `P` = 1+1, `L` = STP. The payout may be `25`, `35 ON OD`, `IRDA` (written as On OD 15; there is no TP rate for STP) or `Rs. 250`
(a fixed rupee amount: written as a **flat card** - the amount goes into the flat-amount column, the rates are 0, and the payout type is the policy's: STP On TP, 1+1 On Net; the payout amount is 70% of the payin amount and the row says so in `remarks`). The payout to the agent is 7 points less. The RTOs of a row are the state minus the RTOs declined for every product (the "PAN India" sheet of the
broker grid: give it with `--broker-grid`, otherwise a built-in October copy is used and the run says so) minus the row's own condition ("HR38/HR51/HR55 are excluded"); the script writes the master group that holds exactly those RTOs.
Next to the upload it writes the sheets `RTO groups to create` / `Make groups to create` of the review workbook (RTO / make groups the master lacks, e.g. `TW_ALL_EX_TVS`) and the sheet `Not written` of the review workbook (the grid rows that gave no row - IRDA on STP, withdrawn - with the reason). The SAOD sheet is not converted yet.

### Go Digit (agency grid): what it does

One password-protected workbook (`Delhi_PAN_India_Agency.xls` - the file is really an encrypted .xlsx). Give the password with `--password`, or the environment variable `GRID_PASSWORD`, or type it when asked (the typing is hidden); the script never stores it. Only the **visible** sheets are read - the hidden ones are older versions.
**Commercial vehicles** (`LATEST CV COMBINED GRID`: cluster | segment | body type | make | CD1 | payout by vehicle age, comprehensive and third party): GCV 4W and 3W, taxis, staff buses, PCV 2W / 3W, MISCD, tractors. **Two-wheelers**: 1+1 with SATP (with / without CPA), 1+5 (only the block titled "1+5"), SAOD. **Private Car**: comprehensive and SAOD for the premium-volume slab of your deal (`PC_VOLUME_SLAB`, now 1-3 Lac), third party by fuel, cc and age.
CD1 is the OD discount the payout is valid for, so it becomes the tariff band: `0-60%` -> 0 to 60.01, `D` = declined = a zero card over the whole range, a pair such as `29%/26%` next to `85%/90%` = two bands. A blank cell is a zero placeholder that never beats a real rate for the same entry; an Excel error cell (`#VALUE!`) counts as a decline. The notes of the sheet `CV Remarks` are applied ("Volvo, Scania to block"; for Truck "Bharat Benz and Mahindra to operate at 70% CD1"). `MISP` cells have no number in the grid: they are written as On OD 17.5 (the October upload's value, `MISP_RATE`) and flagged. School Bus (rates by premium volume) is listed in the review workbook, not written (`WRITE_SCHOOL_BUS`). RTO groups the master lacks (for example HCV Mumbai) are flagged and listed with their RTO codes.
Against the October 2026 upload (47,532 rows) the script reproduces 46,854 rows with the same rates; every other difference is listed in AGENT_CONTEXT.md section 7.8.
**New wordings are read, not flagged (10 Nov 2026).** The discount text of a row (`new, 1 yr cd1 - 60%; 2 yr+ cd1 - 85%`, `0-5 Yr: 0-60% | 6+ Yr: 0-65%`, `0-40% for age 6+ years` ...), the make column of a taxi row (`Maruti & Hyundai age 1+`, `Hyundai: 70%, Maruti 85%`), weight bands, notes that state their own slice and text-typed payouts are read by small grammars; what a grammar does not understand is never guessed (the row keeps the old flag). Every wording read is listed in the review workbook (sheet `CD1 readings`).

**If Go Digit renames or rearranges the sheets** (as in the November 2026 workbook) nothing has to be edited: every table is found by what its header row says and every column by its own header. Where the grid gives two payout columns (`Max CD2` and `Avg CD2`) the script uses `Avg CD2` - your answer of 10 Nov 2026, the setting `PAYOUT_MEASURE` at the top of the script ("max" for the other column, `None` to be asked each run); where the private-car volume slab of your deal is gone it asks which slab to use; a table that cannot be found stops the run with what was expected and what the workbook has (or, if you say so, the file is written without it). Columns the script does not read are listed under WARNING - read that list before you upload.

### SBI General (broker grid): what it does

One workbook: a wide rate sheet (one row per RTO cluster, a block of columns per product: GCV 4W by weight band, GCV 3W, Private Car comprehensive / SAOD / SATP, PCV 3W, taxi, two-wheeler, tractor / harvester, school bus), the High-End enabler sheet, the RTO cluster mapping and the declined make / model lists. SBI moves columns and rows from month to month, so **nothing is found by a cell address**: every block is found by the words in its merged header, and every rule (vehicle age limits, the non-NCB cuts, the zero payout of High-End non-NCB policies, the SATP cut for ages 1-10 ...) is read from the notes under the table by its words. If a note a rule depends on has gone, the script asks you what to do instead of keeping last month's number silently.
**The premium slab is asked every month** (twice: the Private Car slab and the CV slab - the choices are the ones the grid's own notes list; October and November were 1L-3L and 3L-4.99L, one point off every payout of that family; two-wheelers have no slab and are written as printed). Answer on the command line with `--pc-slab 1L-3L --cv-slab 3L-4.99L` (or `none` for the grid as printed).
Comp = 1+1, SATP = STP, SAOD = SAOD; a blank or zero cell is a declined entry and gets a zero card. Notes it cannot apply (the DL +2% for 2.5-3.5 t, the TN-C and WB-K RTO lists, J&K 12-20 t, models 3718 / 3723 ...) are written into the `remarks` column of the rows they may affect; the 'Declined cluster' notes are checked against the grid and every disagreement is flagged on its rows.
Against the October 2026 upload (102,604 rows) the script reproduces 94,176 rows with exactly the same payins and payouts; the differences (the document wins) are listed in CLAUDE.md and AGENT_CONTEXT.md section 7.9.

### ICICI Lombard (CV, Private Car and Two-Wheeler grids): what it does

Give it any of the four workbooks, together or one at a time (`--grid` once per file, or `--cv-grid`, `--pc-grid`, `--tw-grid`, `--tw-new-grid`). Each is recognised by what its sheets say, not by its file name, so a renamed sheet or a moved column does not matter. A run with only some of the workbooks adds `_CV` or `_PC_TW` to the name of the upload. `insurer_vertical` is `ICICI-RB-Offline` on every row (`--vertical` changes it). Each workbook can have its own start date: `--cv-from-date`, `--pc-from-date`, `--tw-from-date`, `--tw-new-from-date` (the new two-wheeler workbook is dated 15 Nov while the others start on 1 Nov).
**Commercial vehicles** (a COMP matrix and an AOTP matrix: one row per RTO cluster, one column per segment): COMP = 1+1 (On Net), AOTP = STP (On TP). A cell is often a small sentence - `OLD:10%,NEW:0%`, `TATA|M&M:55%,OTHERS:35%`, `COMP|NEW|NZD:25%,COMP|NEW|ZD-20%` - and is split into the vehicle ages, makes, policy types and places it names, each with its own rate; `OTHERS` is everything the other parts do not name, and a part at 0% is a decline that the other parts leave out too. Makes are written with the master's make groups (`GCV_4W_ALL_EX_BHARAT_BENZ` ...). A cell the script cannot read with certainty (an unknown word, an Excel error) is **not guessed**: it is listed with the reason in the sheet `Not written` of the review workbook. A cell that says `CC` (Council Call) never becomes a rate: it is listed in the sheet `Council call cells` beside the active Council Call entries, for you to key by hand. `OTHERS:0%` writes no zero card. Notes under the grid that the script cannot apply (the garbage-van rate, the Negative RTO list ...) are written into `remarks` as `Manual Review Required`.
**Private Car** (New = 1+3, SAOD, comprehensive 1+1 (On OD) with and without NCB, Act only = STP) and **Two-Wheeler** (Bike / Scooter, Petrol / Electric: 1+1, SAOD, STP; the new two-wheeler workbook gives 1+5). Rates are written with 8 decimals, like your Private Car upload.
Against your November Private Car + Two-Wheeler upload (3,318 rows) the script reproduces every row with the same rates (only `insurer_vertical` differs, by your decision; 24 Lakshadweep rows are new); against the October CV upload (20,112 rows) it reproduces 19,629 keys, and the differences are listed in AGENT_CONTEXT.md section 7.10.

### IndusInd General (Reliance General) commercial vehicles: what it does

Reliance General and IndusInd General are one insurer; the upload says `INDUSIND GENERAL INSURANCE COMPANY LIMITED`. Give the script the broker TC grid (`Reliance_-_TC_Grid_for_Brokers_Premier-_...xlsx`) and the **two dates**: `--from-date` and `--to-date` are given on every run (the workbook's own title is never read). It converts the commercial-vehicle sheet (`Nov'26`: 73 regions x 31 payout columns -> 157 rows per region); the `Private Car`, `TW Scooter` and `Short Term Taxi` sheets are listed in the review workbook's `Not written` sheet, not converted yet.
Nothing is found by a cell address: the region header (`RTO Region`, usually twice - the parent region and the cluster) and the merged header band above the payouts (make label / product segment / Comp, STP, fuel) are read by their words, and a header word the script does not understand leaves THAT column out with a warning instead of guessing. `new_rto_list` is the region as the grid writes it; a region with no RTO group in the master (Indore in November) is written and flagged. The make groups are the master's own names (`TAXIALL_MAHI_KIA_TOYOTA_MARUTI_MG` ...).
Some rules exist only when their words are in the grid: the model-level qualifier of a make label (`Only Jeeto & Supro`) is flagged on its rows, and a Car Carrier payout above zero is flagged (it would pay every MISCD vehicle). The Guidelines / Contest sheets (School Bus volume slabs, LCV contest, Kaali Pili 40%, hygiene recoveries) are flagged as notes on the rows they name.
Against your own export of the November grid (11,383 rows) it reproduces 11,270 of its 11,304 commercial-vehicle rows exactly; the 34 that differ are the diesel taxi of 17 regions (your file has the next row's value - the script follows the grid and lists them), 157 Indore rows are new, and the export's 79 Private Car rows come from an earlier grid. Everything about this grid - the column-to-rows table, rules, edge cases, differences, how to check a run - is in `INDUSIND_GRID.md`.
The **October 2026** workbook (`Reliance_-_PO_Grid_Broking_Premier_11-Oct_26.xlsx`: sheet `Oct 26`, 72 regions x 30 columns, a top row of group titles, Kaali Pili as a column, deal and Private Car columns inside the sheet, a new `RTO List` sheet) is read by the SAME script - no month-specific code: a header that names its product twice is merged when the texts agree, the new wordings (nil dep, fuel lists, `All Make/All Ages`, `< 5 Years`, `Only deal basis`) are read by their words, Kaali Peeli rows are written with the T&C `Kaali Peeli` and flagged, the M & HCV table stays a flagged note, and a column the script cannot read raises `INCOMPLETE: n of m payout columns were not read` as the first warning. Against your own October export (8,784 rows reproduced, every payin and payout equal) the make groups and T&C it names another way are listed in the review sheet `Other names vs earlier upload`.

### Where the files go

Each run writes into its insurer's folder under `output\` (`liberty`, `magma`, `rsa`, `shriram`, `godigit`, `sbi`, `icici`, `indusind`, `united`; `--out` picks another place): the **upload CSV** `<INSURER>_UPLOAD_<yyyy-mm>.csv` - upload this one, the portal accepts only `.csv` - and, when there is anything to look at, **one review workbook** `<INSURER>_REVIEW_<yyyy-mm>.xlsx` with a sheet each: `Summary`, `RTO groups to create`, `Make groups to create`, `Not written`, `Lowest rate chosen`, `Council call cells` (ICICI), `Manual review notes`. A sheet appears only when it has rows; with nothing to review there is just the CSV. In Colab both files are downloaded.

### Rows that repeat a key (all scripts)

If the grid gives one entry more than once, the same payin twice is one row, and **different payins keep the lowest** - nothing is asked. Every such entry is listed in the sheet `Lowest rate chosen` of the review workbook next to the upload (what was found, what was written) and the end-of-run answers say how many there were.

### Master lists (`master\`)

The two workbooks exported from the portal - the RTO master (RTO name -> RTO codes) and the make / model master (make model name -> make model entries) - are the authority for names. RSA GCV, RSA Private
Car and United check every RTO group and make group they write against them: a name that is not in them is flagged in `remarks` (never replaced by a guess), and United also flags a master group
that does not hold what the circular says (for example the RTO lists of Sub Annexure 3 and 4, or the models of the 7 makes). To refresh them, save the new exports in `master\` (any name with
"rto" + "master", and "make" + "model" + "master"; the newest file wins). The rule is global: Liberty and Magma use it too. They write the master's group name for each geo / cluster
(`MAGMA_APR26_<cluster>`; Liberty `LIBERTY_PC_TW_APR26_` / `LIBERTY_GCV_4W_<7.5T_APR26_` / `LIBERTY_GCV_4W>7.5T_APR26_` + the geo, chosen by the sheet; spaces become `_`), and a geo the master
does not have is flagged in `remarks`, not mapped to another group.

### RSA Private Car: what it does

Comp (1+1), SAOD and New car cells become rows for all five fuels. The top-variant make group gets the cell plus the extra the notes give
(Comp and SAOD only): October's flat "Additional N%", or September's "Top Variant Models Drive" block, whose base-payout lines are 4 points With NCB and
2 points Without NCB (its business-share steps are not in the upload); every other make gets the cell; New car: the top-variant group gets the cell and every other make the flat rate of the note
"Any other make New car payout will be 28% flat". The STP sheet becomes `STP` rows
(Diesel / Non Diesel columns). PO = PI - 7, never below 0; rates are rounded to 2 decimals. The "Top End Variant List" is not a rate table, but the make group
in the portal is built from it - the script prints how many variants each make has, so you can see when the list changed.
Checked against the October upload: all 17,430 rows match. Two differences, both reported every run: the grid has 3 North-East clusters
(ARUNACHAL..., JORHAT..., MEGHALAYA...) that the October upload left out (405 rows), and the October upload had no New rows for other makes, which the
script now writes at the note's flat 28% (1,785 rows; `OTHER_MAKES_NEW = "write"` near the top of the script - set it to `"flag"` to get the October
upload's rows and have the New rows carry the note in `remarks` instead).
Volume-linked payout (VLI) tables are not part of the upload (`pi_vli` = 0) and are listed as standing notes.

### RSA: how the plain-text notes are read

Royal rewrites the notes under each sheet from month to month, so the script reads them instead of using fixed numbers: the Eicher /
other-makes / Nil Dep deductions, the New CV cut (also by weight: "up to 12 ton ... above 12 ton ..."), the discount limits and the extra for
Tata & AL, the effective date, and "the existing grid continues". Each rule is understood in several wordings ("3% lower", "reduced by 3",
"deduction of three"). It prints what it read for every sheet (the `rules from the notes` lines) - check them against the notes.
If a note holds a number the script cannot turn into a rule - or a rule is missing - it does not guess, ask or stop: the rows it may affect
get `Manual Review Required: [the note]` in the `remarks` column (see the next section). Region-specific discount notes (NORTH: a different
maximum discount per cluster and weight) are handled the same way. To see exactly which sentences changed since a good month, before running anything:

```
.\.venv\Scripts\python grid_doctor.py check "<new RSA grid>"
```

## The web application (optional)

`webapp\` is a small Node.js web app that runs these same scripts for you in a browser: choose the insurer, upload the grid, press Run, download the CSV.
When a script would have asked you something, you get a card instead; your answer is remembered and used by itself next time (and the page tells you when it did).
It needs Node 24+; read `webapp\README.md` (set `GRID_APP_PASSWORD`, then `npm install` and `npm start` inside `webapp`).
The scripts themselves are unchanged for you: run from VS Code or a terminal they behave exactly as before.

## Run it in VS Code

1. **File > Open Folder...** and choose `D:\harsh\grid_tools`. (If VS Code offers to install the *Python* extension, say yes.)
2. First time only: double-click `setup.bat`. It builds a private Python environment in `.venv`.
   Run it again if you ever move this folder.
3. Press **Ctrl+Shift+D** (Run and Debug), pick the insurer at the top, press **F5**. All eleven scripts are listed: Liberty, Magma, RSA GCV 4W Comp, RSA Private Car,
   United PDF, Shriram Broker Grid, Shriram Two-Wheeler, Go Digit, SBI General, ICICI Lombard, IndusInd General / Reliance. (If a list from before is still showing, close and reopen the folder in VS Code.)
4. Answer in the **terminal** at the bottom of VS Code:
   - the path of the grid file(s): in the Explorer right-click the file > *Copy Path*, then paste (quotes are fine);
     or put the grids in the `input\` folder first so they are easy to find
   - the earlier upload CSV to compare with (or just press Enter to skip)
   - `from_date` and `to_date` (DD/MM/YYYY)
   - any question the script asks: type the number of an option (`3,4,5` or `3-5` where it says several are allowed),
     `?` for details, `q` to stop. At the end press Enter to write the file.
   - Go Digit only: the workbook's password (the typing is hidden and it is never stored).
5. The CSV appears in `output\`.

### If F5 says "debug type 'debugpy' is not supported"

Either click **Install debugpy Extension** in that message and press F5 again, or skip the debugger completely:
**Terminal > Run Task...** and pick the matching **Run ...** task, one per insurer (needs no extension), or use the terminal commands below.

## Or run it from a terminal

```
cd D:\harsh\grid_tools
.\.venv\Scripts\python liberty_grid_to_config.py
.\.venv\Scripts\python magma_grid_to_config.py
.\.venv\Scripts\python rsa_gcv4w_comp_to_config.py
.\.venv\Scripts\python rsa_pvtcar_to_config.py
.\.venv\Scripts\python united_commission_to_config.py
.\.venv\Scripts\python shriram_broker_grid_to_config.py
.\.venv\Scripts\python shriram_2w_to_config.py
.\.venv\Scripts\python godigit_agency_to_config.py
.\.venv\Scripts\python icici_grid_to_config.py
.\.venv\Scripts\python indusind_grid_to_config.py
```

Everything can also be given up front (nothing is asked unless something needs a decision):

```
.\.venv\Scripts\python liberty_grid_to_config.py --grid "D:\harsh\grids\...\Sep26_Branch_Grid_v3.xlsx" --from-date 01/09/2026 --to-date 31/03/2030 --out output\LIBERTY.csv
.\.venv\Scripts\python magma_grid_to_config.py --grid "<comprehensive>.xlsx" --satp "<satp>.xlsb" --from-date 01/10/2026 --to-date 31/03/2030
.\.venv\Scripts\python rsa_gcv4w_comp_to_config.py --grid "<RSA GCV comp grid>.xlsx" --from-date 01/10/2026 --to-date 31/03/2030 --ref "<last upload>.csv"
.\.venv\Scripts\python rsa_pvtcar_to_config.py --grid "<RSA Pvt Car grid>.xlsx" --from-date 01/10/2026 --to-date 31/03/2030 --ref "<last upload>.csv"
.\.venv\Scripts\python shriram_broker_grid_to_config.py --grid "<Shriram broker grid>.xlsx" [--addendum "<addendum>.xlsx"] --from-date 01/10/2026 --to-date 31/03/2030 --ref "<last upload>.csv"
.\.venv\Scripts\python shriram_2w_to_config.py --grid "<Shriram 2W grid>.xlsx" [--broker-grid "<Shriram broker grid>.xlsx"] --from-date 01/10/2026 --to-date 31/03/2030 --ref "<last upload>.csv"
.\.venv\Scripts\python godigit_agency_to_config.py --grid "<Delhi_PAN_India_Agency.xls>" --password "<password>" --from-date 15/10/2026 --to-date 31/03/2030 --ref "<last upload>.csv"
.\.venv\Scripts\python icici_grid_to_config.py --grid "<CV grid>.xlsx" --grid "<Private Car grid>.xlsx" --grid "<TW grid>.xlsx" --grid "<new TW grid>.xlsx" --from-date 01/11/2026 --to-date 31/03/2030 [--tw-new-from-date 15/11/2026] --ref "<last upload>.csv"
.\.venv\Scripts\python indusind_grid_to_config.py --grid "<Reliance TC grid>.xlsx" --from-date 01/11/2026 --to-date 31/03/2030 --ref "<last upload>.csv"
```

RSA only: `--prefix ROYAL_SEPT26_` sets the start of every `new_rto_list` (the portal's RTO groups still carry the September prefix).

Add `--no-prompt` to never be asked anything (it stops with the question as the error instead).

## Notes the scripts cannot apply: the `remarks` column

Insurers write rules in plain text (notes, guidelines, remarks) that no script can always turn into logic. In **every** script, such a note is
never guessed at, never asked about and never stops the run. Instead the rows it may affect carry, in the `remarks` column of the upload,

```
Manual Review Required: [the note text]
```

(several notes on one row are joined with ` | `). Filter that column in the CSV, take those notes up with the insurer, and map those rates by hand
afterwards. The script prints `MANUAL REVIEW - N rows ...` with each note and how many rows it flagged. Which rows: RSA flags the clusters and make groups
a note names (a state brings in all its clusters; a place it cannot place, such as NCR, or none named, means the whole sheet) and the weight bands a New CV
reduction does not clearly cover; Liberty and Magma flag the sheet (Liberty) or grid (Magma) the text came from. Texts already known to need no action
are listed in `STANDING_NOTES` near the top of each script's build section - add a phrase there once you have decided a note needs no action.
The comparison with an earlier upload ignores `remarks`.

## If something goes wrong

- A plain `STOPPED - ...` line says what is wrong; fix the file or answer differently and run again.
- An unexpected problem is saved to `output\*_error_*.txt` with the technical details.

## When an insurer changes the grid format (the fixing agent, included in the Claude Pro plan)

Most changes are handled by the script itself: it notices the difference and asks you a question. For the rest, use Claude Code
pointed at this folder - it is part of the Pro plan, and nothing here calls the paid API.

1. Open Claude Code **in this folder**: in the Claude desktop app start a Code session with the folder `D:\harsh\grid_tools`,
   or in a terminal: `cd D:\harsh\grid_tools` then `claude`.
2. Type `/fix-grid` followed by the grid path(s) and, if you have it, the error line, for example
   `/fix-grid "D:\harsh\grids\NOV26\liberty\Nov26 Branch Grid.xlsx" it stops with "slab ... appears twice"`.
3. It reads `lessons_learned.md`, checks what changed in the file (`grid_doctor.py`), reproduces the stop, makes the smallest
   fix, re-runs the known-good grids (`check_regressions.py` must say `ALL OK`) and shows you the proof.
   It stops and asks you if it would have to guess a rate or a column's meaning.
4. When you approve, it remembers the new layout and writes the lesson down so the same surprise never costs time twice.

Before the first month of a new grid, you can also check for changes without Claude (no usage at all):

```
.\.venv\Scripts\python grid_doctor.py check "<new grid>"
```

It prints exactly which sheets, headers or text values moved compared with the closest remembered month.

Notes: it uses your Claude *subscription*; if the environment variable `ANTHROPIC_API_KEY` is set on this PC, Claude Code
bills that key instead, so leave it unset. Run `python check_regressions.py` any time you have edited a script by hand.

## What the agent remembers about each insurance company (one skill per insurer)

Every insurer has its own skill in `.claude\skills\<insurer>-grid\SKILL.md`: `shriram-grid`, `sbi-grid`, `icici-grid`, `liberty-grid`, `magma-grid`, `rsa-grid` (both Royal Sundaram scripts), `united-grid`, `godigit-grid`, `indusind-grid`.
The agent loads it together with `grid-converter` whenever you send that insurer's grid or its export, so it already knows, before it asks you anything:

- **Status** - what was built, what you accepted, how many rows;
- **Rules you have set (do not re-ask)** - for example "STP rows always have NA in NCB and ZD", "United 40% annexure: Package 1+1 only", "SBI two-wheelers have no slab";
- **Traps seen** - what went wrong once and how to spot it;
- **Open items** - what is still waiting for you or for a new master export;
- **Monthly procedure** - the steps for that insurer's grid.

After every grid the agent adds what it learned to that insurer's skill (and to `AGENT_CONTEXT.md`, the long record). When you send a corrected export with your hand fixes, it compares it with its own output, finds every change you made and turns it into a rule.
A new script cannot be added without its skill: `python check_regressions.py` and `python tests\skills_tests.py` fail if one is missing.

### One dedicated reference file per insurer

Next to each skill sits that insurer's full reference in this folder: `SHRIRAM_GRID.md`, `SBI_GRID.md`, `ICICI_GRID.md`, `LIBERTY_GRID.md`, `MAGMA_GRID.md`, `RSA_GRID.md` (both Royal Sundaram scripts), `UNITED_GRID.md`, `GODIGIT_GRID.md`, `INDUSIND_GRID.md`.
Each one describes that insurer's grid, how it becomes upload rows, the names, the notes, the edge cases, how to check a run, your decisions and what is still open. The skill is the short version the agent reads first; the `.md` is the detail.

**These files are rule-based.** They never name a place in a grid by a cell address, a column range, a row number or a column letter - a place is named by its labels (the words of a sheet's title, of a header, of a row's label) and by the rule that finds it, so the text stays true when the insurer re-lays the sheet out.
`python .claude\skills\grid-converter\scripts\cellref_lint.py` lists any slip; `python check_regressions.py` and `python tests\insurer_docs_tests.py` fail on one, and the agent's `learn.py` refuses to save one. A new script cannot be added without its `.md` either.

