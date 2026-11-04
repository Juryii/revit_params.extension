"""Draw the two ribbon icons from editable vector primitives (requires Pillow)."""
from pathlib import Path
from xml.sax.saxutils import quoteattr
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
BUTTONS = {
    'purchase': ROOT / 'Params.tab/Крепления.panel/Расчет закупки.pushbutton',
    'parameters': ROOT / 'Params.tab/INK.panel/Edit_params.pushbutton',
}


def primitives(name, dark=False):
    ink = '#dbeafe' if dark else '#24466b'
    blue = '#60bfff' if dark else '#168bd1'
    orange = '#ffbd69' if dark else '#e88a24'
    if name == 'purchase':
        return [
            ('rect', dict(x=2, y=3, width=27, height=5, rx=1.5, fill=blue)),
            ('rect', dict(x=2, y=11, width=16, height=5, rx=1.5, fill=ink)),
            ('rect', dict(x=20, y=11, width=9, height=5, rx=1.5, fill=blue)),
            ('line', dict(x1=20, y1=18, x2=29, y2=29, stroke=orange)),
            ('line', dict(x1=29, y1=18, x2=20, y2=29, stroke=orange)),
            ('circle', dict(cx=19, cy=28, r=2.6, fill='none', stroke=orange)),
            ('circle', dict(cx=29, cy=28, r=2.6, fill='none', stroke=orange)),
            ('line', dict(x1=3, y1=24, x2=12, y2=24, stroke=ink)),
            ('line', dict(x1=3, y1=21, x2=3, y2=27, stroke=ink)),
            ('line', dict(x1=12, y1=21, x2=12, y2=27, stroke=ink)),
        ]
    return [
        ('line', dict(x1=4, y1=y, x2=28, y2=y, stroke=ink)) for y in (7, 16, 25)
    ] + [
        ('rect', dict(x=x, y=y-3.5, width=6, height=7, rx=1.6, fill=color))
        for x, y, color in [(8, 7, blue), (19, 16, orange), (12, 25, blue)]
    ]


def draw(name, dark=False):
    shapes = primitives(name, dark)
    scale = 8
    image = Image.new('RGBA', (32*scale, 32*scale))
    canvas = ImageDraw.Draw(image)
    svg = ['<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" '
           'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">']
    for kind, attrs in shapes:
        svg.append('<{} {}/>'.format(kind, ' '.join('{}={}'.format(k, quoteattr(str(v))) for k, v in attrs.items())))
        fill = attrs.get('fill')
        fill = None if fill == 'none' else fill
        stroke = attrs.get('stroke')
        if kind == 'rect':
            x, y, w, h = [attrs[k]*scale for k in ('x', 'y', 'width', 'height')]
            canvas.rounded_rectangle((x, y, x+w, y+h), radius=attrs['rx']*scale, fill=fill)
        elif kind == 'circle':
            x, y, r = [attrs[k]*scale for k in ('cx', 'cy', 'r')]
            canvas.ellipse((x-r, y-r, x+r, y+r), fill=fill, outline=stroke, width=2*scale)
        else:
            x1, y1, x2, y2 = [attrs[k]*scale for k in ('x1', 'y1', 'x2', 'y2')]
            canvas.line((x1, y1, x2, y2), fill=stroke, width=2*scale)
            for x, y in [(x1, y1), (x2, y2)]:
                canvas.ellipse((x-scale, y-scale, x+scale, y+scale), fill=stroke)
    svg.append('</svg>')
    stem = 'icon.dark' if dark else 'icon'
    folder = BUTTONS[name]
    (folder / (stem+'.svg')).write_text('\n'.join(svg)+'\n', encoding='utf-8')
    image.resize((32, 32), Image.Resampling.LANCZOS).save(folder / (stem+'.png'))


if __name__ == '__main__':
    for button in BUTTONS:
        for dark in (False, True):
            draw(button, dark)
