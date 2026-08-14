# Voynich Dual-Track AI Edition

An auditable research and creative-writing pilot for the Voynich Manuscript (Beinecke MS 408).

> [!CAUTION]
> No generally accepted decipherment of the Voynich Manuscript exists. This project does **not** claim to have translated it. The analytical edition reports observable structure without assigning English meanings; the companion English text is a deterministic, illustration-informed **creative reconstruction**, not a translation.

## Download the pilot editions

- [`Voynich_Pilot_Evidence_Based_Analytical_Edition.docx`](editions/Voynich_Pilot_Evidence_Based_Analytical_Edition.docx) — the responsible scholarly track: facsimile, EVA transcription, token statistics, positional observations, and explicit confidence limits.
- [`Voynich_Pilot_Evidence_Based_Analytical_Edition.pdf`](editions/Voynich_Pilot_Evidence_Based_Analytical_Edition.pdf) — fixed-layout rendering of the analytical edition.
- [`Voynich_Pilot_Speculative_English_Reconstruction.docx`](editions/Voynich_Pilot_Speculative_English_Reconstruction.docx) — the experimental track: a readable English reconstruction, prominently labeled as unverified fiction.
- [`Voynich_Pilot_Speculative_English_Reconstruction.pdf`](editions/Voynich_Pilot_Speculative_English_Reconstruction.pdf) — fixed-layout rendering of the creative reconstruction.

Release-file checksums are listed in [`SHA256SUMS`](SHA256SUMS).

## What is in the pilot

The fixed sample contains 12 surfaces spanning the major conventional sections of the manuscript:

| Folio | Conventional section | Purpose in the sample |
|---|---|---|
| f1r | text-only/control | Tests the method without illustration-led semantic cues |
| f3r | herbal A | Early herbal page |
| f31r | herbal B | Different Currier language and scribal hand |
| f67r1 | astronomical | Circular and radial text |
| f70r2 | cosmological | Diagram labels and concentric text |
| f70v2 | zodiac/Pisces | Short labels around figures and stars |
| f75r | biological/balneological | Pools, figures, and running text |
| fRos | cosmological foldout | Large nine-rosette diagram |
| f88r | pharmaceutical | Plant parts, labels, and paragraphs |
| f99r | pharmaceutical | Inventory-like labels and containers |
| f103r | star-marked entries | Opening of the final text section |
| f116r | star-marked entries | Corpus-boundary test near the manuscript's end |

The generated dataset contains 535 loci and 3,380 normalized token groups. Every row preserves its folio and locus identifier.

## The two tracks

### 1. Evidence-based analytical edition

This track preserves the source EVA, supplies a readable EVA view, counts forms against the complete ZL corpus, and describes only observable features such as paragraph position, label status, repetition, and common form families. Its semantic confidence is explicitly zero.

### 2. Speculative English reconstruction

This track uses the illustration category and locus type to select from a frozen English phrase inventory. SHA-256-based selection makes the wording deterministic. The procedure creates a consistent reading experience, but it does not derive English words from Voynichese and is not evidence of decipherment.

Keeping the tracks separate prevents an attractive hypothesis from silently becoming a claimed result.

## Reproduce the build

Requirements: Python 3.11 or newer, `python-docx`, and Pillow.

```bash
python -m pip install -r requirements.txt
python src/fetch_zl.py
python src/build_pilot.py \
  --zl sources/ZL3b-n.txt \
  --image-dir assets/folios \
  --repo .
python -m unittest discover -s tests
```

The fetch script verifies the pinned ZL3b-n SHA-256 hash before saving it. The source file itself is not committed because it is a modern scholarly transcription with its own rights status.

## Sources and status

- Manuscript and institutional description: [Yale Beinecke Library](https://beinecke.library.yale.edu/beinecke/collections/beinecke-cipher-voynich-manuscript)
- Digital manuscript record: [Yale Digital Collections, Beinecke MS 408](https://collections.library.yale.edu/catalog/2002046)
- Transliteration documentation: [René Zandbergen, “Transliteration of the Voynich MS”](https://www.voynich.nu/transcr.html)
- Pinned input: [Zandbergen–Landini EVA, IVTFF 2.0, ZL3b-n (13 May 2025)](https://www.voynich.nu/data/ZL3b-n.txt), SHA-256 `bf5b6d4ac1e3a51b1847a9c388318d609020441ccd56984c901c32b09beccafc`

Yale describes the manuscript as undeciphered. Computational and linguistic work has produced useful tests and competing models, but not a broadly validated English translation. See [`docs/limitations.md`](docs/limitations.md) and [`docs/methodology.md`](docs/methodology.md).

This repository does not claim to be the first use of AI on the manuscript. Public projects already make AI-assisted decipherment claims, including [kamb-code/Voynich](https://github.com/kamb-code/Voynich) and [scott-schechter/voynich-decoded](https://github.com/scott-schechter/voynich-decoded). Their claims are listed as related work, not endorsed here.

## Project status

Version `0.1.0` is a deliberately small pilot. The rule set is frozen before any expansion. A full-manuscript edition should proceed only after independent review of the data model, labels, and separation between evidence and invention.

## AI system disclosure

This edition was prepared with OpenAI Codex using **GPT-5.6 Sol** at **max reasoning effort**. The model produced the editorial framing, analytical descriptions, deterministic creative reconstruction, code, and publication files under the limitations stated above. Model disclosure does not convert the speculative English reconstruction into a verified translation.

## Rights and attribution

- Code written for this project: MIT License; see [`LICENSE`](LICENSE).
- Original editorial prose, tables, and speculative reconstruction: CC BY 4.0; see [`LICENSE-CONTENT.md`](LICENSE-CONTENT.md).
- Manuscript images and modern source transcription retain their respective source status and are not relicensed by this repository.

Prepared with OpenAI Codex - GPT-5.6 Sol (max reasoning).
