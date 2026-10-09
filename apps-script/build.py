#!/usr/bin/env python3
"""루트 index.html(단일 파일)을 Apps Script용 4개 파일(index/style/app1/app2)로 나눈다.
붙여넣기가 잘리는 것을 막기 위해 파일당 줄 수를 줄이는 용도. 사용: python3 apps-script/build.py"""
import re, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(__file__).resolve().parent
s = (root / 'index.html').read_text(encoding='utf-8').replace('<!--END-OF-INDEX-->', '').rstrip() + '\n'
style = re.search(r'<style>.*?</style>', s, re.S).group(0)
script = re.search(r'<script>.*?</script>', s, re.S).group(0)
body = script[len('<script>'):-len('</script>')].strip('\n')
lines = body.split('\n')
cut = next(i for i, l in enumerate(lines) if l.startswith('/* ---------- 일별 입력 패널'))
app1 = '<script>\n' + '\n'.join(lines[:cut]) + '\n</script>\n'
app2 = '<script>\n' + '\n'.join(lines[cut:]) + '\n</script>\n'
index = s.replace(style, "<?!= include_('style') ?>").replace(script, "<?!= include_('app1') ?>\n<?!= include_('app2') ?>")
# 끝 표식: HtmlService가 HTML 주석을 지우므로 주석이 아닌 형태로 넣는다
tail = {
    'index': None,
    'style': '<style>#END-OF-style{display:none}</style>\n',
    'app1': "<script>'END-OF-app1';</script>\n",
    'app2': "<script>'END-OF-app2';</script>\n",
}
index = index.replace('</body>', '<div id="END-OF-index" hidden></div>\n</body>')
for name, text in [('index', index), ('style', style + '\n'), ('app1', app1), ('app2', app2)]:
    (out / f'{name}.html').write_text(text + (tail[name] or ''), encoding='utf-8')
