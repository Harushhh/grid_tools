"""Copy the shared question / safeguard block of one script into all the others.

    python sync_shared.py liberty_grid_to_config.py

Every `*_to_config.py` carries the same block between its  '# %% --- QUESTIONS'  and  '# %% --- READING THE ...'  headings
(Asker, the cell reader, ManualReview, the CSV writer ...). Edit it in ONE script, then run this so the others get the identical text;
check_regressions.py fails when the copies differ.
"""
import glob
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
BLOCK = re.compile(r"^# %% -+ QUESTIONS\n.*?(?=^# %% -+ READING THE )", re.S | re.M)


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    source = os.path.join(HERE, sys.argv[1])
    block = BLOCK.search(open(source, encoding="utf-8").read())
    if not block:
        sys.exit(f"no shared block found in {sys.argv[1]}")
    for path in sorted(glob.glob(os.path.join(HERE, "*_to_config.py"))):
        if os.path.samefile(path, source):
            continue
        text = open(path, encoding="utf-8").read()
        new, n = BLOCK.subn(lambda m: block.group(0), text, count=1)
        if n != 1:
            print(f"SKIPPED {os.path.basename(path)}: no shared block found")
            continue
        if new != text:
            open(path, "w", encoding="utf-8", newline="\n").write(new)
            print(f"updated {os.path.basename(path)}")
        else:
            print(f"already identical: {os.path.basename(path)}")


if __name__ == "__main__":
    main()
