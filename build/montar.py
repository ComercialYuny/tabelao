"""Monta o index.html do Tabelão a partir de dados.json + template.html.
Uso (na pasta build/):  python3 montar.py   -> grava ../index.html
"""
import json, os
AQUI = os.path.dirname(os.path.abspath(__file__))
d = json.load(open(os.path.join(AQUI, 'dados.json'), encoding='utf-8'))
t = open(os.path.join(AQUI, 'template.html'), encoding='utf-8').read()
logo = d.pop('logo')
js = json.dumps(d, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
out = t.replace('__LOGO__', logo).replace('__DATA__', js)
dest = os.path.join(AQUI, '..', 'index.html')
open(dest, 'w', encoding='utf-8').write(out)
print('index.html gravado:', round(len(out) / 1e6, 2), 'MB ·', d['mes'], '· atualizado', d['atualizado'], '· status VIMOB', d['status'])
