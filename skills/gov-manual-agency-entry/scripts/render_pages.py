"""Render PDF pages of a Government Manual to PNG so they can be read as images.

Usage: python render_pages.py <pdf> <out_dir> <dpi> <pdf_page> [<pdf_page> ...]
  110-120 dpi is enough for unit lists and most charts; use 150 for dense fold-out charts.
Then open each PNG with the Read tool and decide parents from indentation, boxes, and lines.
"""
import os, sys
import pymupdf

pdf, out, dpi = sys.argv[1], sys.argv[2], int(sys.argv[3])
os.makedirs(out, exist_ok=True)
doc = pymupdf.open(pdf)
stem = os.path.splitext(os.path.basename(pdf))[0].replace(' ', '_')
for p in map(int, sys.argv[4:]):
    pix = doc[p - 1].get_pixmap(dpi=dpi)
    fn = os.path.join(out, f'{stem}_p{p}.png')
    pix.save(fn)
    print(fn, pix.width, 'x', pix.height)
