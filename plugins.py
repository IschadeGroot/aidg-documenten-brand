#!/usr/bin/env python3
"""
De AI van AIDG Documenten in de editors-inhoud zetten.

Gebruik:  python3 plugins.py <checkout-van-Euro-Office/plugins> <common-map>

Euro-Office Desktop levert de editors zonder plugins. Online (AIDG Werkplek)
zit het AI-tabblad er wel in; dit zet dezelfde plugin erbij, uit
Euro-Office/plugins (Apache-2.0), plus:
- de plugin-bibliotheek v1 (de plugins laden ../v1/plugins.js);
- het koppelscript (koppeling/aidg-koppeling.js) vóór de eigen scripts van
  de AI, zodat AIDG vanzelf de aanbieder is (via de AIDG-app, met het eigen
  account en gekozen model). Zie de app-repo, src/main/aidg-documenten.

Welke plugins mee moeten staat in PLUGINS; de map-naam in de editors is de
GUID uit de config.json van de plugin (zoals de editor ze verwacht).
"""
from __future__ import annotations

import json
import re
import shutil
import sys
from pathlib import Path

HIER = Path(__file__).resolve().parent
PLUGINS = ['ai']
KOPPEL = {'ai'}  # plugins die het koppelscript krijgen


def main() -> int:
    bron, common = Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve()
    doelmap = common / 'editors' / 'sdkjs-plugins'
    doelmap.mkdir(parents=True, exist_ok=True)

    v1 = bron / 'sdkjs-plugins' / 'v1'
    if not v1.is_dir():
        print(f'FOUT: geen v1 in {bron}')
        return 1
    shutil.copytree(v1, doelmap / 'v1', dirs_exist_ok=True)
    print('v1 geplaatst')

    for naam in PLUGINS:
        p = bron / 'sdkjs-plugins' / 'content' / naam
        cfg = json.loads((p / 'config.json').read_text('utf-8-sig'))
        guid = re.sub(r'^asc\.', '', cfg['guid'])
        doel = doelmap / guid
        shutil.copytree(p, doel, dirs_exist_ok=True, ignore=shutil.ignore_patterns('.dev', '.git*', 'node_modules'))
        # De bron laadt de plugin-bibliotheek van internet (onlyoffice.github.io);
        # in de desktop-app komt die uit de meegeleverde v1-map (offline, en er
        # gaat niets naar een externe server).
        for html in doel.rglob('*.html'):
            t = html.read_text('utf-8')
            nieuw = re.sub(r'https?://onlyoffice\.github\.io/sdkjs-plugins/v1/', '../v1/', t)
            if nieuw != t:
                html.write_text(nieuw, 'utf-8')
        if naam in KOPPEL:
            shutil.copyfile(HIER / 'koppeling' / 'aidg-koppeling.js', doel / 'aidg-koppeling.js')
            html = doel / 'index.html'
            tekst = html.read_text('utf-8')
            if 'aidg-koppeling.js' not in tekst:
                i = tekst.find('<script')
                if i < 0:
                    print(f'FOUT: geen <script> in {html}')
                    return 1
                tekst = tekst[:i] + '<script type="text/javascript" src="aidg-koppeling.js"></script>\n    ' + tekst[i:]
                html.write_text(tekst, 'utf-8')
        print(f'plugin {naam} -> {guid}' + (' (met AIDG-koppeling)' if naam in KOPPEL else ''))
    return 0


if __name__ == '__main__':
    sys.exit(main())
