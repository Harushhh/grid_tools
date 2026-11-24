"""Render the pages of a PDF to PNG files so they can be LOOKED AT (by a person, or by an agent that reads images).

    render_pages.py <file.pdf> [--out DIR] [--pages 1,3-5] [--resolution 130] [--crop-images]

Why: scanned circulars (United's cover letter every month, the whole October 2026 circular) have no usable text layer, and the cover letter holds rules that change the
rows. --crop-images crops each page to the picture it holds (scanned tables are often a small landscape picture in the middle of a portrait page) and renders THAT, which is
much easier to read; use a higher --resolution (200+) for small print. Writes <DIR>/<pdf name>_p<N>.png (default DIR: next to the pdf, in 'pages_<pdf name>') and prints the paths
plus, per page, how many pictures / text characters / tables it found (0 text and 1 picture = a scan).
Needs pdfplumber (in the grid_tools .venv). Standard library otherwise."""
import argparse
import os
import sys


def page_numbers(spec, last):
    if not spec:
        return list(range(1, last + 1))
    out = []
    for part in spec.split(","):
        a, _, b = part.strip().partition("-")
        out += list(range(int(a), int(b or a) + 1))
    return [n for n in out if 1 <= n <= last]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf")
    ap.add_argument("--out", help="output folder")
    ap.add_argument("--pages", help="e.g. 1,3-5 (default: all)")
    ap.add_argument("--resolution", type=int, default=130)
    ap.add_argument("--crop-images", action="store_true", help="crop each page to its (first) picture before rendering")
    a = ap.parse_args(argv)
    try:
        import pdfplumber
    except ImportError:
        sys.exit("pdfplumber is not installed - run it with the grid_tools venv: .venv\\Scripts\\python.exe")
    stem = os.path.splitext(os.path.basename(a.pdf))[0]
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.pdf)), f"pages_{stem}")
    os.makedirs(out, exist_ok=True)
    with pdfplumber.open(a.pdf) as pdf:
        for n in page_numbers(a.pages, len(pdf.pages)):
            page = pdf.pages[n - 1]
            facts = f"{len(page.images)} picture(s), {len(page.extract_text() or '')} text characters, {len(page.find_tables())} table(s)"
            target = page
            if a.crop_images and page.images:
                im = page.images[0]
                box = (max(im["x0"] - 2, 0), max(im["top"] - 2, 0), min(im["x1"] + 2, page.width), min(im["bottom"] + 2, page.height))
                if box[2] > box[0] and box[3] > box[1]:
                    target = page.crop(box)
            path = os.path.join(out, f"{stem}_p{n}.png")
            target.to_image(resolution=a.resolution).save(path)
            print(f"page {n}: {facts} -> {path}")


if __name__ == "__main__":
    main()
