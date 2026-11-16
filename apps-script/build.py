#!/usr/bin/env python3
"""루트 index.html(단일 파일)을 Apps Script용 여러 파일(index/style*/app*)로 나눈다.
붙여넣기가 잘리는 것을 막기 위해 파일당 분량을 줄이는 용도. 사용: python3 apps-script/build.py"""
import re, pathlib
here = pathlib.Path(__file__).resolve().parent
s = (here.parent / 'index.html').read_text(encoding='utf-8')
for old in list(here.glob('style*.html')) + list(here.glob('app*.html')):
    old.unlink()
parts = {}
def split(tag, prefix, text):
    n = 0
    def sub(m):
        nonlocal n
        n += 1
        name = f'{prefix}{n}'
        parts[name] = m.group(0)
        return f"<?!= include_('{name}') ?>"
    return re.sub(rf'<{tag}>.*?</{tag}>', sub, text, flags=re.S)
index = split('script', 'app', split('style', 'style', s))
index = index.replace('</body>', '<div id="END-OF-index" hidden></div>\n</body>')
# 끝 표식: HtmlService가 HTML 주석을 지우므로 주석이 아닌 형태로 넣는다
for name, text in parts.items():
    tail = f"<style>#END-OF-{name}{{display:none}}</style>\n" if name.startswith('style') else f"<script>'END-OF-{name}';</script>\n"
    (here / f'{name}.html').write_text(text.rstrip('\n') + '\n' + tail, encoding='utf-8')
(here / 'index.html').write_text(index, encoding='utf-8')
names = ['index'] + list(parts)
code = (here / 'Code.gs').read_text(encoding='utf-8')
code = re.sub(r"const PARTS = \[.*?\];", "const PARTS = [" + ', '.join(f"'{n}'" for n in names) + "];", code, count=1)
(here / 'Code.gs').write_text(code, encoding='utf-8')
print(names)
