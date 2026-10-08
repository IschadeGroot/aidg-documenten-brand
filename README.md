# AIDG Documenten — merklaag en build

AIDG Documenten is het kantoorprogramma van AIDG voor de pc: **Euro-Office
DesktopEditors** (de Europese office die ook online op AIDG Werkplek draait),
met het AIDG-merk en de AI van AIDG erin.

Deze repo bevat **alleen de AIDG-laag**. De broncode van het programma komt
ongewijzigd van [Euro-Office/DesktopEditors](https://github.com/Euro-Office/DesktopEditors);
de build legt deze laag erover en maakt een Windows-installer.

## Wat er in de laag zit

| Map/bestand | Wat |
|---|---|
| `merk/` | AIDG-pictogram (`app-icon.png`, `app-icon.ico`) en woordmerken (licht/donker, lang/kort) |
| `merk.py` | zoekt de merkbestanden van Euro-Office (alles met `-eo` in de naam, de logo's in de editors) en schrijft er de AIDG-versie voor, op precies dezelfde maat; zet de namen in de installer op "AIDG Documenten" |
| `plugins.py` + `koppeling/` | zet de AI-plugin uit [Euro-Office/plugins](https://github.com/Euro-Office/plugins) erin (het AI-tabblad, zoals online), met de plugin-bibliotheek lokaal (niets van internet) en de AIDG-koppeling: de AI gebruikt vanzelf AIDG via de AIDG-app |
| `euro-office-ref.txt` | welke Euro-Office-commit gebouwd wordt |
| `plugins-ref.txt` | welke commit van Euro-Office/plugins |

## Bouwen

- **Automatisch:** elke maandag kijkt `Euro-Office-updates` of Euro-Office een
  nieuwere commit heeft die bij hen zelf groen bouwde. Zo ja: vastzetten,
  bouwen, release.
- **Met de hand:** Actions → *Bouw AIDG Documenten* → Run workflow (eventueel
  met een andere commit).
- Een wijziging in de laag (logo's, scripts) naar `main` start ook een build.

Uitkomst: een release met `AIDG-Documenten-<versie>-x64.exe` (installer, voor
iedereen op de pc) en `SHA256SUMS.txt`. De AIDG-app installeert precies die
versie en controleert de SHA-256.

## Logo vervangen

Zet nieuwe bestanden in `merk/` met dezelfde namen (pictogram als PNG van
1024×1024 én als ICO; woordmerken als SVG) en push naar `main`. De rest
(alle maten, licht/donker) gaat vanzelf.

## Licentie

Euro-Office DesktopEditors is AGPL-3.0; de AI-plugin is Apache-2.0. De
volledige bron van elke release is de vermelde Euro-Office-commit plus deze
repo op de vermelde commit. Het Over-venster van het programma blijft
vermelden dat het op Euro-Office (en ONLYOFFICE) gebaseerd is.
