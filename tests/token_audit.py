"""Function-level wording audit (18 Nov 2026, the user: "at some places there is some make or other conditions which is not applied, also it is not mentioned anywhere").

A converter reads a free-text cell with a function (a label -> a kind, an age text -> a range, a note -> rules).  A word of the text that the function does not use is a condition that is not
applied: unless the function says so (a flag, a refusal), nobody ever sees it.  This tool finds those words WITHOUT reading the function: it deletes the text's words one at a time and
asks the function again; a word whose deletion changes nothing is IGNORED.

    audit(func, texts)  -> {text: [ignored words]}        func(text) -> anything comparable (a dict, a tuple, a string)
    tally(result)       -> {word: how many texts ignore it}     (a word ignored EVERYWHERE is a filler; a word ignored in one text only is the interesting one)

The function should return what it makes of the text INCLUDING its flags / refusals: a word the function flags is not silent.  A word that is part of a string the function echoes
(a flag that quotes the whole text) changes the echo when deleted, so it counts as read - which is right: it is mentioned."""
import re

WORD = re.compile(r"[A-Za-z0-9][A-Za-z0-9+./&%'\-]*|[<>=]+")


def words(text):
    return [(m.start(), m.end()) for m in WORD.finditer(text)]


def audit(func, texts, drop_filler=()):
    """{text: [ignored words]} for the texts that have any; `drop_filler` = words not worth reporting (articles ...)"""
    out = {}
    skip = {w.lower() for w in drop_filler}
    for t in texts:
        try:
            base = repr(func(t))
        except Exception as e:
            out[t] = [f"<function raised {type(e).__name__}>"]
            continue
        ignored = []
        for s, e in words(t):
            w = t[s:e]
            if w.lower() in skip:
                continue
            t2 = t[:s] + " " + t[e:]
            try:
                same = repr(func(t2)) == base
            except Exception:
                same = False
            if same:
                ignored.append(w)
        if ignored:
            out[t] = ignored
    return out


def tally(result):
    n = {}
    for ws in result.values():
        for w in set(x.lower() for x in ws):
            n[w] = n.get(w, 0) + 1
    return dict(sorted(n.items(), key=lambda kv: -kv[1]))


def show(result, limit=60, width=150):
    for t, ws in list(result.items())[:limit]:
        print(f"  {t[:width]!r}\n      ignored: {ws}")
    if len(result) > limit:
        print(f"  ... {len(result) - limit} more")
