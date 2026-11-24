"""The insurer skills (.claude/skills/<insurer>-grid/SKILL.md): the knowledge the agent loads for one insurance company (the user, 4 Nov 2026: "creating a skill insurance company wise, for our agent to understand").
Every converter script has its insurer's skill, each skill has the same sections and a trigger description, grid-converter points at all of them, and `learn.py skill` updates them.
Run:  .venv\\Scripts\\python.exe tests\\skills_tests.py"""
import glob
import importlib.util
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

ROOT = helpers.ROOT
SKILLS = os.path.join(ROOT, ".claude", "skills")
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


def front(path):
    """(name, description) of a SKILL.md's frontmatter, or (None, None)"""
    text = open(path, encoding="utf-8").read()
    m = re.match(r"---\r?\nname: (?P<name>[^\r\n]+)\r?\ndescription: (?P<desc>[^\r\n]+)\r?\n---", text)
    return (m.group("name").strip(), m.group("desc").strip()) if m else (None, None)


scripts = sorted(os.path.basename(p) for p in glob.glob(os.path.join(ROOT, "*_to_config.py")))
insurers = sorted({s.split("_")[0] for s in scripts})
check("there are converter scripts to cover", len(scripts) >= 9 and {"shriram", "sbi", "rsa", "liberty", "magma", "united", "godigit", "icici"} <= set(insurers), insurers)

missing = [i for i in insurers if not os.path.isfile(os.path.join(SKILLS, f"{i}-grid", "SKILL.md"))]
check("every converter script has its insurer's skill (.claude/skills/<insurer>-grid/SKILL.md; the insurer = the first word of the script's file name)", not missing,
      f"missing: {missing} - create the skill when you add a script (grid-converter Job C) and `learn.py skill` keeps it current")

SECTIONS = ("## Status", "## Rules the user", "## Traps seen", "## Open items", "## Monthly procedure")
for i in insurers:
    path = os.path.join(SKILLS, f"{i}-grid", "SKILL.md")
    if not os.path.isfile(path):
        continue
    name, desc = front(path)
    text = open(path, encoding="utf-8").read().lower()
    check(f"{i}-grid: frontmatter name = the folder name, and a description that says when to load it (the insurer's name, the word 'grid-converter')",
          name == f"{i}-grid" and desc and i in desc.lower() and "grid-converter" in desc.lower() and len(desc) > 150, (name, (desc or "")[:80]))
    check(f"{i}-grid: has the sections Status / Rules the user has set / Traps seen / Open items / Monthly procedure (the places `learn.py skill` writes to)",
          all(any(ln.startswith(s.lower()) for ln in text.splitlines()) for s in SECTIONS), [s for s in SECTIONS if not any(ln.startswith(s.lower()) for ln in text.splitlines())])
    check(f"{i}-grid: names its script(s) and says where the detail lives (AGENT_CONTEXT.md)", all(s in text for s in [x.lower() for x in scripts if x.split("_")[0] == i]) and "agent_context.md" in text,
          [x for x in scripts if x.split("_")[0] == i and x.lower() not in text])

conv = open(os.path.join(SKILLS, "grid-converter", "SKILL.md"), encoding="utf-8").read()
check("grid-converter points at every insurer skill (so the agent loads both)", all(f"{i}-grid" in conv for i in insurers), [i for i in insurers if f"{i}-grid" not in conv])
check("grid-converter tells the agent to update the insurer's skill after a run (`learn.py skill`)", "learn.py skill" in conv)

# learn.py skill, on a temporary copy of the skills. The dates typed below are in the past ON PURPOSE: a --date that is ignored falls back to today's date,
# and a test that typed today's date would pass by coincidence (it did, on 5 Nov 2026: the mutation 'ignore --date' survived).
spec = importlib.util.spec_from_file_location("learn", os.path.join(SKILLS, "grid-converter", "scripts", "learn.py"))
learn = importlib.util.module_from_spec(spec)
spec.loader.exec_module(learn)
tmp = tempfile.mkdtemp(prefix="skills_test_")
shutil.copytree(SKILLS, os.path.join(tmp, ".claude", "skills"), ignore=shutil.ignore_patterns("__pycache__"))
path = learn.add_skill_note(tmp, "shriram", "status", "the   November  output was accepted", "17 Feb 2026")
text = open(path, encoding="utf-8").read()
status = text.split("## Status")[1].split("\n## ")[0]
check("learn.py skill adds a dated bullet at the END of the matching section (whitespace tidied), in the right skill, before the next section", status.rstrip().endswith("- (17 Feb 2026) the November output was accepted")
      and path.endswith(os.path.join("shriram-grid", "SKILL.md")), status[-120:])
learn.add_skill_note(tmp, "shriram-grid", "traps", "a trap", "17 Feb 2026")
learn.add_skill_note(tmp, "sbi", "open", "an open item", "17 Feb 2026")
sbi = open(os.path.join(tmp, ".claude", "skills", "sbi-grid", "SKILL.md"), encoding="utf-8").read()
check("... 'shriram-grid' and 'shriram' both name the skill; each section keeps its own bullet", "- (17 Feb 2026) a trap" in open(path, encoding="utf-8").read().split("## Traps seen")[1].split("\n## ")[0]
      and sbi.rstrip().endswith("- (17 Feb 2026) an open item"), sbi[-120:])
os.makedirs(os.path.join(tmp, ".claude", "skills", "zz-grid"))
open(os.path.join(tmp, ".claude", "skills", "zz-grid", "SKILL.md"), "w", encoding="utf-8").write("---\nname: zz-grid\ndescription: t\n---\n# ZZ\n\n## Status\n- first\n")
learn.add_skill_note(tmp, "zz", "rules", "a rule", "17 Feb 2026")
zz = open(os.path.join(tmp, ".claude", "skills", "zz-grid", "SKILL.md"), encoding="utf-8").read()
check("a section that is not there yet is created at the end of the skill", zz.rstrip().endswith("## Rules the user has set (do not re-ask)\n- (17 Feb 2026) a rule") and zz.count("## ") == 2, zz)
try:
    learn.add_skill_note(tmp, "nosuchinsurer", "status", "x")
    check("an insurer without a skill stops with the list of the skills that exist", False)
except SystemExit as e:
    check("an insurer without a skill stops with the list of the skills that exist", "nosuchinsurer-grid" in str(e) and "shriram" in str(e) and "sbi" in str(e), str(e)[:200])
# the command line the agent really types (argument names, --root, the choices of --section), on the same temporary copy
LEARN = os.path.join(SKILLS, "grid-converter", "scripts", "learn.py")
p = subprocess.run([sys.executable, "-I", "-X", "utf8", LEARN, "--root", tmp, "skill", "--insurer", "united", "--section", "traps", "--text", "a cli trap", "--date", "17 Feb 2026"],
                   capture_output=True, text=True, encoding="utf-8")
united = open(os.path.join(tmp, ".claude", "skills", "united-grid", "SKILL.md"), encoding="utf-8").read()
check("the command line works: `learn.py --root <copy> skill --insurer united --section traps --text ... --date ...` adds the bullet to the Traps section",
      p.returncode == 0 and "added to" in p.stdout and "- (17 Feb 2026) a cli trap" in united.split("## Traps seen")[1].split("\n## ")[0], (p.returncode, p.stdout[-200:], p.stderr[-200:]))
p = subprocess.run([sys.executable, "-I", "-X", "utf8", LEARN, "--root", tmp, "skill", "--insurer", "united", "--section", "nonsense", "--text", "a refused bullet"],
                   capture_output=True, text=True, encoding="utf-8")
united = open(os.path.join(tmp, ".claude", "skills", "united-grid", "SKILL.md"), encoding="utf-8").read()
check("... and a section that does not exist is refused, nothing is written", p.returncode != 0 and "a refused bullet" not in united, (p.returncode, p.stderr[-200:]))
shutil.rmtree(tmp, ignore_errors=True)

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
