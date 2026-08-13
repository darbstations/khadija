# -*- coding: utf-8 -*-
"""دمج التشغيل مع تحليل الموقع — صفحة تفاعلية للخمس محطات"""
import json, pathlib
D = pathlib.Path(__file__).parent
data = json.load(open(D/'five.json', encoding='utf-8'))
head = open(D/'five_head.html', encoding='utf-8').read()
body = open(D/'five_body.html', encoding='utf-8').read()
body = body.replace('__DATA__', json.dumps(data, ensure_ascii=False, separators=(',', ':')))
out = head + "\n" + body
open('/home/user/khadija/docs/خمس-محطات-تفاعلي.html', 'w', encoding='utf-8').write(out)
print(f'كتب docs/خمس-محطات-تفاعلي.html — {len(out):,} حرف')
