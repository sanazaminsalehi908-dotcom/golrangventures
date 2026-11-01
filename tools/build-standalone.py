"""Build a single self-contained HTML file of the kids' idea form to send around.

The published artifact loads its images from img/ and stores ideas through Claude. A copy opened
anywhere else has neither, so this inlines every image as a data URI and wraps the page in a full
document; the page then runs in its preview mode (everything works, nothing is stored).

Usage: python3 tools/build-standalone.py [OUT]   (default forms/dist/golrang-kids-ideas.html)
"""
import base64
import json
import pathlib
import re
import sys

root = pathlib.Path(__file__).resolve().parent.parent
src = (root / 'forms/golrang-kids-ideas.html').read_text(encoding='utf8')
out = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else root / 'forms/dist/golrang-kids-ideas.html'

imgs = {}
for f in sorted((root / 'forms/img').glob('*.webp')):
    imgs[f.name] = 'data:image/webp;base64,' + base64.b64encode(f.read_bytes()).decode()

# static references in CSS and markup
page = re.sub(r'img/([\w.-]+\.webp)', lambda m: imgs[m.group(1)], src)
# dynamic references in the script go through asset(), which reads window.__IMG
page = page.replace('<script>', '<script>window.__IMG = ' + json.dumps(imgs) + ';</script>\n<script>', 1)

doc = ('<!doctype html><html lang="fa" dir="rtl"><head><meta charset="utf-8">'
       '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
       '<style>:root{padding-top:env(safe-area-inset-top);padding-bottom:env(safe-area-inset-bottom)}'
       'body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>'
       '</head><body>\n' + page + '\n</body></html>\n')
out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(doc, encoding='utf8')
print(out, f'{len(doc.encode()) / 1024:.0f} KB')
