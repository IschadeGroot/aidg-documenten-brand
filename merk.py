#!/usr/bin/env python3
"""
AIDG-huisstijl over een checkout van Euro-Office DesktopEditors leggen.

Gebruik:  python3 merk.py <checkout> <uitvoermap>

Schrijft in <uitvoermap> een boom met dezelfde paden als de checkout; die
boom wordt in de build over de checkout uitgepakt (zelfde mechanisme als de
merklaag van Euro-Office zelf). Er wordt niets in de checkout aangepast.

Hoe het de bestanden vindt (geen vaste lijst, zodat nieuwe merkbestanden in
een volgende Euro-Office-versie vanzelf meegaan):
- alle bestanden met "-eo" in de naam onder de merkmappen van desktop-apps
  (de Euro-Office-varianten van pictogrammen en logo's);
- de logo's in de koppen van de editors (web-apps, img/header en img/about).

Per bestand:
- .ico                → het AIDG-programmapictogram;
- pictogram (app-icon, icon, desktopeditors) → AIDG-pictogram op dezelfde maat;
- logo/woordmerk      → AIDG-woordmerk, licht of donker naar gelang het
                        origineel (gemeten aan de kleuren van het origineel);
- .svg blijft .svg (zelfde breedte/hoogte), .png blijft .png (zelfde pixels).

Wat niet te herkennen is, wordt gemeld en overgeslagen (de build faalt dan
niet, maar het staat in het log).
"""
from __future__ import annotations

import io
import os
import re
import shutil
import sys
from pathlib import Path

from PIL import Image

HIER = Path(__file__).resolve().parent
MERK = HIER / 'merk'

ZOEK_EO = [
    'desktop-apps/win-linux/res/icons',
    'desktop-apps/win-linux/res/img',
    'desktop-apps/common/loginpage/res/img',
    'desktop-apps/win-linux/extras',
]
ZOEK_EDITOR = [
    'web-apps/apps/common/main/resources/img/header',
    'web-apps/apps/common/main/resources/img/about',
]
EXT = ('.svg', '.png', '.ico')


def svg_maat(tekst: str) -> tuple[float, float] | None:
    b = re.search(r'\bwidth="([\d.]+)', tekst)
    h = re.search(r'\bheight="([\d.]+)', tekst)
    if b and h:
        return float(b.group(1)), float(h.group(1))
    vb = re.search(r'viewBox="[\d.\-]+\s+[\d.\-]+\s+([\d.]+)\s+([\d.]+)"', tekst)
    return (float(vb.group(1)), float(vb.group(2))) if vb else None


def helderheid_svg(tekst: str) -> float | None:
    """Gemiddelde helderheid (0-1) van de vulkleuren, of None als er geen zijn."""
    kleuren = re.findall(r'(?:fill|stop-color|color)\s*[:=]\s*"?\s*#([0-9a-fA-F]{3,6})\b', tekst)
    waarden = []
    for k in kleuren:
        if len(k) == 3:
            k = ''.join(c * 2 for c in k)
        if len(k) != 6:
            continue
        r, g, b = (int(k[i : i + 2], 16) / 255 for i in (0, 2, 4))
        waarden.append(0.2126 * r + 0.7152 * g + 0.0722 * b)
    if re.search(r'fill\s*[:=]\s*"?\s*(white|#fff\b)', tekst, re.I):
        waarden.append(1.0)
    return sum(waarden) / len(waarden) if waarden else None


def helderheid_png(pad: Path) -> float | None:
    im = Image.open(pad).convert('RGBA')
    im.thumbnail((64, 64))
    som = n = 0.0
    for r, g, b, a in getattr(im, "get_flattened_data", im.getdata)():
        if a > 128:
            som += (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255
            n += 1
    return som / n if n else None


def soort(pad: Path) -> str | None:
    naam = pad.name.lower()
    if pad.suffix.lower() == '.ico':
        return 'ico'
    if 'splash' in naam:
        return 'splash'
    if re.search(r'app-icon|appicon|desktopeditors|^icon', naam):
        return 'pictogram'
    if 'logo' in naam:
        return 'logo'
    return None


def is_licht(pad: Path) -> bool:
    """Is het origineel licht van kleur (dan staat het op een donkere achtergrond)?"""
    if pad.suffix.lower() == '.svg':
        h = helderheid_svg(pad.read_text('utf-8', errors='replace'))
    else:
        h = helderheid_png(pad)
    if h is None:
        return 'dark' in pad.name.lower() or 'white' in pad.name.lower()
    return h > 0.6


def woordmerk_svg(licht: bool, breed: bool) -> str:
    naam = ('woordmerk' if breed else 'kort') + ('-licht' if licht else '-donker') + '.svg'
    return (MERK / naam).read_text('utf-8')


def maak_svg(origineel: Path, svg: str) -> str:
    """Het AIDG-svg op de maat van het origineel (breedte/hoogte), passend en links/midden."""
    maat = svg_maat(origineel.read_text('utf-8', errors='replace'))
    if not maat:
        return svg
    b, h = maat
    svg = re.sub(r'\s(width|height)="[^"]*"', '', svg, count=2)
    return svg.replace(
        '<svg ',
        f'<svg width="{b:g}" height="{h:g}" preserveAspectRatio="xMinYMid meet" ',
        1,
    )


def render_svg(svg: str, b: int, h: int) -> Image.Image:
    import cairosvg  # alleen nodig voor logo's als png

    png = cairosvg.svg2png(bytestring=svg.encode('utf-8'), output_height=h)
    im = Image.open(io.BytesIO(png)).convert('RGBA')
    if im.width > b:
        im = im.resize((b, round(im.height * b / im.width)), Image.LANCZOS)
    doek = Image.new('RGBA', (b, h), (0, 0, 0, 0))
    doek.paste(im, (0, (h - im.height) // 2), im)
    return doek


def pictogram_png(b: int, h: int) -> Image.Image:
    bron = Image.open(MERK / 'app-icon.png').convert('RGBA')
    zijde = min(b, h)
    klein = bron.resize((zijde, zijde), Image.LANCZOS)
    doek = Image.new('RGBA', (b, h), (0, 0, 0, 0))
    doek.paste(klein, ((b - zijde) // 2, (h - zijde) // 2), klein)
    return doek


def pictogram_svg(origineel: Path) -> str:
    import base64

    data = base64.b64encode((MERK / 'app-icon.png').read_bytes()).decode('ascii')
    maat = svg_maat(origineel.read_text('utf-8', errors='replace')) or (64, 64)
    b, h = maat
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
        f'width="{b:g}" height="{h:g}" viewBox="0 0 {b:g} {h:g}">'
        f'<image width="{b:g}" height="{h:g}" preserveAspectRatio="xMidYMid meet" '
        f'xlink:href="data:image/png;base64,{data}"/></svg>'
    )


def splash_svg(b: float = 500, h: float = 250) -> str:
    """Opstartscherm: wit, grijze rand, AIDG-pictogram links, naam ernaast."""
    import base64

    data = base64.b64encode((MERK / 'app-icon.png').read_bytes()).decode('ascii')
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
        f'width="{b:g}" height="{h:g}" viewBox="0 0 500 250">'
        '<rect width="500" height="250" fill="#fff"/>'
        '<rect x="0.5" y="0.5" width="499" height="249" fill="none" stroke="#ACACAC"/>'
        f'<image x="40" y="65" width="120" height="120" xlink:href="data:image/png;base64,{data}"/>'
        # textLength: past altijd, ook met een breder lettertype dan Segoe UI
        '<text x="182" y="137" font-family="Segoe UI, Arial, Liberation Sans, DejaVu Sans, sans-serif" '
        'font-size="32" fill="#1f2a33" textLength="290" lengthAdjust="spacingAndGlyphs">'
        '<tspan font-weight="700">AIDG</tspan> Documenten</text>'
        '</svg>'
    )


# Wat de gebruiker in de installer en Windows ziet. Intern (installatiemap,
# register, AppId) blijft het Euro-Office: zo blijven hun geteste paden en
# updates gewoon werken.
INSTALLER_NAMEN = {
    'sAppName': '"AIDG Documenten"',
    'sAppPublisher': '"AIDG"',
    'sAppPublisherURL': '"https://aidg.nl/"',
    'sAppSupportURL': '"https://aidg.nl/"',
    'sAppIconName': '"AIDG Documenten"',
    'ASSC_APP_NAME': '"AIDG Documenten"',
    'ASCC_REG_REGISTERED_APP_NAME': '"AIDG Documenten"',
    'ASSOC_APP_FRIENDLY_NAME': '"AIDG Documenten"',
}


def installer_namen(root: Path, uitmap: Path) -> int:
    bron = root / 'desktop-apps' / 'package' / 'inno' / 'defines.iss'
    if not bron.is_file():
        print('let op: geen defines.iss gevonden; installernamen niet aangepast')
        return 0
    tekst = bron.read_text('utf-8-sig')
    n = 0
    for naam, waarde in INSTALLER_NAMEN.items():
        tekst, k = re.subn(rf'^(#define\s+{naam}\s+).*$', lambda m: m.group(1) + waarde, tekst, flags=re.M)
        n += k
        if not k:
            print(f'let op: {naam} niet in defines.iss')
    doel = uitmap / bron.relative_to(root)
    doel.parent.mkdir(parents=True, exist_ok=True)
    doel.write_text(chr(0xFEFF) + tekst, 'utf-8')
    print(f'installer: {n} namen naar AIDG Documenten')
    return n


def doelen(root: Path) -> list[Path]:
    uit: list[Path] = []
    for rel in ZOEK_EO:
        d = root / rel
        if d.is_dir():
            uit += [p for p in d.rglob('*') if p.is_file() and p.suffix.lower() in EXT and '-eo' in p.stem.lower()]
    for rel in ZOEK_EDITOR:
        d = root / rel
        if d.is_dir():
            uit += [p for p in d.glob('*') if p.is_file() and p.suffix.lower() in ('.svg', '.png') and 'logo' in p.name.lower()]
    return sorted(set(uit))


def main() -> int:
    root, uitmap = Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve()
    gedaan = overgeslagen = 0
    for pad in doelen(root):
        rel = pad.relative_to(root)
        doel = uitmap / rel
        doel.parent.mkdir(parents=True, exist_ok=True)
        s = soort(pad)
        ext = pad.suffix.lower()
        try:
            if s == 'ico':
                shutil.copyfile(MERK / 'app-icon.ico', doel)
            elif s == 'pictogram' and ext == '.svg':
                doel.write_text(pictogram_svg(pad), 'utf-8')
            elif s == 'pictogram' and ext == '.png':
                b, h = Image.open(pad).size
                pictogram_png(b, h).save(doel)
            elif s == 'splash' and ext == '.svg':
                maat = svg_maat(pad.read_text('utf-8', errors='replace')) or (500, 250)
                doel.write_text(splash_svg(*maat), 'utf-8')
            elif s == 'splash' and ext == '.png':
                b, h = Image.open(pad).size
                render_svg(splash_svg(), b, h).save(doel)
            elif s == 'logo' and ext == '.svg' and (lambda m: m and m[0] / max(m[1], 1) < 1.5)(
                svg_maat(pad.read_text('utf-8', errors='replace'))
            ):
                # vierkant "logo" (bv. het pictogram op het startscherm): het AIDG-pictogram
                doel.write_text(pictogram_svg(pad), 'utf-8')
            elif s == 'logo':
                licht = is_licht(pad)
                if ext == '.svg':
                    maat = svg_maat(pad.read_text('utf-8', errors='replace')) or (100, 20)
                    svg = woordmerk_svg(licht, breed=maat[0] / max(maat[1], 1) > 4.5)
                    doel.write_text(maak_svg(pad, svg), 'utf-8')
                else:
                    b, h = Image.open(pad).size
                    svg = woordmerk_svg(licht, breed=b / max(h, 1) > 4.5)
                    render_svg(svg, b, h).save(doel)
            else:
                print(f'overgeslagen (onbekend soort): {rel}')
                overgeslagen += 1
                continue
            print(f'AIDG: {rel}')
            gedaan += 1
        except Exception as e:  # noqa: BLE001 — één bestand mag de rest niet tegenhouden
            print(f'FOUT bij {rel}: {e}')
            overgeslagen += 1
    installer_namen(root, uitmap)
    print(f'{gedaan} bestanden in AIDG-huisstijl, {overgeslagen} overgeslagen')
    return 0 if gedaan else 1


if __name__ == '__main__':
    sys.exit(main())
