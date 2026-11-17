#!/usr/bin/env python3
"""Генерирует брендированные копии демки из таблицы лидов.
Использование: python3 tools/generate.py leads.xlsx  -> папки /<slug>/index.html и tools/links.json
"""
import sys, re, json, os, openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = open(os.path.join(ROOT, 'index.html'), encoding='utf-8').read()
SITE = 'https://vaditik23.github.io/himchistka-demo/'
SKIP = {'italclean', 'expert-chistoty', 'chisty-vopros', 'tools'}

TR = dict(zip('абвгдеёжзийклмнопрстуфхцчшщъыьэюя',
  ['a','b','v','g','d','e','e','zh','z','i','y','k','l','m','n','o','p','r','s','t','u','f','h','c','ch','sh','sch','','y','','e','yu','ya']))
def slugify(s):
    s = s.lower().split('/')[0].split('(')[0]
    s = ''.join(TR.get(c, c if re.match(r'[a-z0-9]', c) else '-') for c in s)
    return re.sub(r'-+', '-', s).strip('-')

def js(s): return s.replace('\\', '\\\\').replace("'", "\\'")
def h(s): return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def plural(n, f):
    n = abs(n) % 100; n1 = n % 10
    if 10 < n < 20: return f[2]
    if 1 < n1 < 5: return f[1]
    if n1 == 1: return f[0]
    return f[2]

def rating_perk(r):
    m = re.search(r'2ГИС (\d\.\d) \((\d+)', r or '')
    if not m: return None
    val, cnt = float(m.group(1)), int(m.group(2))
    if val < 4.7 or cnt < 20: return None
    return f'⭐ {str(val).replace(".", ",")} на 2ГИС · {cnt} {plural(cnt, ["оценка","оценки","оценок"])}'

def build(name, city, profile, rating, key):
    p = profile.lower()
    carpets = 'ковр' in p and 'мебел' not in p and 'диван' not in p
    delivery = 'доставк' in p or 'забор' in p
    leather = 'кож' in p
    city_disp = re.sub(r'\s*\(.*?\)', '', city).replace(' / ', ', ')
    city_main = city_disp.split(',')[0].strip()
    tagline = profile[0].upper() + profile[1:]
    s = BASE
    def rep(a, b, count=1):
        nonlocal s
        assert a in s, a[:60]
        s = s.replace(a, b, count)
    rep('<title>Чисто&amp;Свежо — заявка на химчистку мебели (демо)</title>',
        f'<title>{h(name)} — {"стирка ковров" if carpets else "химчистка мебели"}, {h(city_main)}</title>')
    rep("name:'Чисто&Свежо',city:'Москва',tagline:'Выездная химчистка мебели и ковров'",
        f"name:'{js(name)}',city:'{js(city_disp)}',tagline:'{js(tagline)}'")
    rep("tgChat:'@chisto_svezho_bot'", f"tgChat:'@{key.replace('-', '_')}_bot'")
    rep("const KEY='chs-demo-v1'", f"const KEY='demo-{key}-v1'")
    rep("{id:'in',n:'В пределах города',fee:0}", f"{{id:'in',n:'{js(city_main)}',fee:0}}")
    if carpets:
        rep('<h1>Узнайте цену химчистки за 1 минуту</h1>', '<h1>Узнайте цену стирки ковра за 1 минуту</h1>')
        rep('<p>Выберите, что чистим, и сразу увидите стоимость. Без звонков и «ой, это не по прайсу».</p>',
            '<p>Укажите размер ковра и сразу увидите стоимость. Без звонков и ожидания ответа.</p>')
        # ковёр первым в списке
        i = s.index("  {id:'carpet'"); j = s.index('\n', i) + 1; line = s[i:j]; s = s[:i] + s[j:]
        k = s.index("  {id:'sofa'"); s = s[:k] + line + s[k:]
    else:
        rep('<p>Выберите, что чистим, и сразу увидите стоимость. Без звонков и «ой, это не по прайсу».</p>',
            '<p>Выберите, что чистим, и сразу увидите примерную стоимость. Без звонков и ожидания ответа.</p>')
    perks = []
    rp = rating_perk(rating)
    if rp: perks.append(rp)
    if delivery: perks.append('🚚 Забор и доставка')
    if leather: perks.append('🛋 Чистим кожу')
    perks += ['🧮 Цена за 1 минуту', '📷 Расчёт по фото', '🗓 Удобное время']
    i = s.index('<div class="perks">'); j = s.index('</div>', i)
    s = s[:i] + '<div class="perks">' + ''.join(f'<span class="perk">{h(x)}</span>' for x in perks) + s[j:]
    # отзывы -> как это работает (без выдуманных отзывов)
    i = s.index('<div class="steptitle"><h2 style="font-size:16px">Что говорят клиенты</h2></div>')
    j = s.index('</div></div>`;', i) + len('</div></div>')
    s = s[:i] + ('<div class="steptitle"><h2 style="font-size:16px">Как это работает</h2></div><div class="reviews">'
        '<div class="rev"><b>1. Выберите, что чистить</b><br>Цена считается сразу, со всеми доплатами и скидками. Никаких сюрпризов.</div>'
        '<div class="rev"><b>2. Добавьте фото</b><br>Мастер заранее увидит загрязнения и подберёт химию.</div>'
        '<div class="rev"><b>3. Выберите время</b><br>Свободные окна видно сразу. Мастер подтвердит заявку в удобном мессенджере.</div></div>') + s[j:]
    rep('Химия, пятновыводители и выезд по городу уже включены. Оплата после работы, картой или переводом.',
        'Химия и пятновыводители уже включены. Точную цену мастер подтвердит до начала работы.')
    rep('<span>Оплата после работы, когда всё понравится</span>', '<span>Оплата после работы</span>')
    streets = ['ул. Ленина, 104','ул. Советская, 12','ул. Мира, 78','ул. Гагарина, 9','ул. Садовая, 21','ул. Пушкина, 3','ул. Победы, 50','ул. Московская, 26']
    olds = re.findall(r"\{name:'[^']*',phone:'[^']*',contact:'[^']*',addr:'([^']*)'", s)
    for o, n in zip(olds, streets): s = s.replace(f"addr:'{o}'", f"addr:'{n}'", 1)
    assert 'Москва' not in s and 'Кунцево' not in s
    return s

def main(xlsx):
    ws = openpyxl.load_workbook(xlsx)['Лиды']
    hdr = [c.value for c in ws[1]]
    col = {n: i for i, n in enumerate(hdr)}
    used, links = set(SKIP), {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[col['Название']]: continue
        num, name, city = row[col['#']], str(row[col['Название']]).strip(), str(row[col['Город']]).strip()
        disp = re.sub(r'\s*\(.*?\)', '', name).split(' / ')[0].strip() or name
        slug = slugify(disp)
        if slug in used: slug = slug + '-' + slugify(city)
        used.add(slug)
        os.makedirs(os.path.join(ROOT, slug), exist_ok=True)
        open(os.path.join(ROOT, slug, 'index.html'), 'w', encoding='utf-8').write(
            build(disp, city, str(row[col['Профиль']] or ''), str(row[col['Рейтинг / отзывы']] or ''), slug))
        links[num] = SITE + slug + '/'
    json.dump(links, open(os.path.join(ROOT, 'tools', 'links.json'), 'w'), ensure_ascii=False, indent=1)
    print(len(links), 'demos')

if __name__ == '__main__':
    main(sys.argv[1])
