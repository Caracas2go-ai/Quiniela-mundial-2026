#!/usr/bin/env python3
"""
build_site.py — Genera quiniela-app.html: la web app en UN solo archivo
autocontenido (CSS + logo + data incrustados) para previsualizar/compartir.
El index.html de producción usa styles.css + data.json + logo.png externos.
"""
import sys, os, json, base64

base = sys.argv[1] if len(sys.argv) > 1 else 'quiniela-web'
base = os.path.join(os.path.dirname(os.path.abspath(__file__)), base)

html = open(os.path.join(base, 'index.html'), encoding='utf-8').read()
sb = os.path.join(base, 'score-bg.jpg')
if os.path.isfile(sb):
    _b = base64.b64encode(open(sb, 'rb').read()).decode('ascii')
    html = html.replace("url('score-bg.jpg')", "url('data:image/jpeg;base64," + _b + "')")
data = json.load(open(os.path.join(base, 'data.json'), encoding='utf-8'))
data['css'] = open(os.path.join(base, 'styles.css'), encoding='utf-8').read()
lp = os.path.join(base, 'logo.png')
if os.path.isfile(lp):
    data['logo'] = base64.b64encode(open(lp, 'rb').read()).decode('ascii')

inject = '<script>window.__DATA__=' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';</script>\n</head>'
out = html.replace('</head>', inject, 1)
dst = os.path.join(base, 'quiniela-app.html')
open(dst, 'w', encoding='utf-8').write(out)
print(f'[build_site] {os.path.getsize(dst)/1024:.0f} KB -> {dst}')
