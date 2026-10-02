# -*- coding: utf-8 -*-
"""تدقيق آلي قبل التسليم: python3 audit.py source.txt main.docx [reading.docx]"""
import re, sys, zipfile
src = open(sys.argv[1], encoding='utf-8').read()
res = []
def chk(name, ok, detail=''):
    res.append((name, ok, detail)); print(('PASS ' if ok else 'FAIL ') + name + (' - ' + detail if detail else ''))
# ----- المصدر
chk('توازن < و >', src.count('<') == src.count('>'), '%d/%d' % (src.count('<'), src.count('>')))
nested = any('<' in m for m in re.findall(r'<([^>]*)>', src))
chk('لا حواشي متداخلة', not nested)
chk('توازن @@( و )@@', src.count('@@(') == src.count(')@@'))
chk('لا ملاحظة صفراء معلقة', '@@(' not in src)
bad = [i + 1 for i, l in enumerate(src.split('\n')) if l.count('"') % 2]
chk('توازن علامات التنصيص في كل فقرة', not bad, str(bad))
chk('لا مسافة قبل علامة الحاشية', not re.search(r'\s<', src))
fns = re.findall(r'<([^>]*)>', src)
chk('عدد الحواشي', True, str(len(fns)))
# كتاب بلا جزء وصفحة
BOOKS = ['فتاوى السبكي', 'تكملة المجموع', 'الإبهاج', 'المستصفى', 'البرهان', 'الإحكام للآمدي', 'شرح تنقيح الفصول', 'المعتمد',
         'البحر المحيط', 'تيسير التحرير', 'كشف الأسرار', 'أحكام القرآن للكيا', 'الكشاف', 'الانتصاف المطبوع بهامش', 'تفسير ابن كثير', 'درء تعارض',
         'مجموع الفتاوى', 'مفتاح دار السعادة', 'الفصل في الملل', 'الإشارات الإلهية', 'التحبير', 'طبقات الشافعية',
         'مقاييس اللغة', 'لسان العرب', 'المحكم', 'المخصص', 'مختار الصحاح', 'كشف الظنون']
missing = []
for f in fns:
    for b in BOOKS:
        for m in re.finditer(re.escape(b), f):
            tail = f[m.end(): m.end() + 80]
            if not re.search(r'\((\d+/ ?[\d، \-]+|ص: ?[\d\- ]+)\)', tail.split('؛')[0]) and not re.search(r'\(\d+/ \d', f[m.start():m.start()+140]) and not re.search(r'ص: \d', f[m.start():m.start()+140]):
                missing.append((b, f[:60]))
chk('كل كتاب مذكور في حاشية له جزء وصفحة', not missing, str(missing[:5]))
# ----- ملفات docx
for path in sys.argv[2:]:
    z = zipfile.ZipFile(path)
    doc = z.read('word/document.xml').decode('utf-8')
    texts = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', doc, flags=re.S))
    chk('[%s] لا أرقام لاتينية في المتن' % path, not re.search(r'[0-9]', texts), re.findall(r'.{10}[0-9].{10}', texts)[:2].__str__())
    if 'footnotes.xml' in ' '.join(z.namelist()):
        fx = z.read('word/footnotes.xml').decode('utf-8')
        ftexts = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', fx, flags=re.S))
        chk('[%s] لا أرقام لاتينية في الحواشي' % path, not re.search(r'[0-9]', ftexts))
        nref = doc.count('<w:footnoteReference')
        chk('[%s] عدد مراجع الحواشي = عدد الحواشي' % path, nref == len(fns), '%d' % nref)
    else:
        chk('[%s] انتفاء footnoteReference (الأصل بعلامات < >)' % path, '<w:footnoteReference' not in doc)
        chk('[%s] علامات الحواشي في المتن' % path, texts.count('&lt;') + texts.count('<') >= len(fns))
    chk('[%s] لا تظليل أصفر' % path, 'w:highlight' not in doc)
    chk('[%s] لا شرطة طويلة' % path, '—' not in texts and '–' not in texts)
    chk('[%s] مسافة غير منقسمة داخل العزو' % path, '/ ' in texts)
    chk('[%s] خط Traditional Arabic وحجم ١٨' % path, 'Traditional Arabic' in z.read('word/styles.xml').decode('utf-8') and 'w:sz w:val="36"' in z.read('word/styles.xml').decode('utf-8'))
print('\nFAILS:', [r[0] for r in res if not r[1]])
