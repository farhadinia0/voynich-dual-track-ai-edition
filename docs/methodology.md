# Methodology

## Research question

Can a transparent pipeline turn a modern EVA transcription into two useful outputs without confusing structural analysis with semantic translation?

The answer tested here is procedural, not cryptanalytic: preserve evidence in one track, place imaginative interpretation in another, and make every transformation inspectable.

## Fixed pilot sample

The 12 selected surfaces were chosen before generation. Together they cover text-only, herbal, astronomical, zodiac, biological/balneological, cosmological, pharmaceutical, and star-marked material. Folios f3r and f31r deliberately cross the conventional Currier A/B division; f1r removes the visual-semantic crutch; f116r tests the far end of the corpus.

The fixed order is recorded in `data/pilot_ruleset.json` and `data/build_manifest.json`.

## Input and parsing

The input is Zandbergen–Landini EVA transliteration file ZL3b-n, IVTFF 2.0, version 3b (13 May 2025). Its expected SHA-256 is:

```text
bf5b6d4ac1e3a51b1847a9c388318d609020441ccd56984c901c32b09beccafc
```

The parser retains:

- folio identifier;
- locus identifier and position code;
- locus type (paragraph, label, circular/radial text, and so forth);
- source EVA string;
- a display-oriented EVA string;
- normalized token groups.

Normalization exists only for counting. It removes IVTFF editorial markup, resolves common alternatives conservatively, and extracts lowercase EVA-like alphabetic groups. The untouched source string remains in every row so another researcher can reject or replace the normalization.

## Analytical track

Counts are computed over all 5,385 loci parsed from the pinned ZL file, not only the pilot. For each pilot locus, the analytical edition reports:

- line position and locus type;
- normalized token count;
- within-line repetition;
- up to three recurrent or productive-ending forms;
- an explicit statement that no English lexical value is established.

The output deliberately avoids plant identifications, medical diagnoses, ingredient mappings, phonetic values, and plaintext claims. Illustration descriptions are contextual metadata, not translations.

## Speculative track

The creative edition uses fixed phrase inventories for each conventional section. A stable SHA-256 digest of the rule-set version, section, semantic slot, and local EVA material selects phrases. Therefore identical input under the same rule set produces identical prose.

This is not a cipher key. There is no learned mapping from EVA glyphs or tokens to English words. The procedure demonstrates how easily coherent prose can be generated from layout and imagery alone—which is precisely why coherence cannot validate a decipherment.

## Confidence model

Two confidence domains are kept separate:

| Domain | Pilot status |
|---|---|
| Folio/locus identity | High, subject to the cited transcription |
| EVA display and normalized token extraction | Moderate to high; editorial choices are inspectable |
| Positional and frequency observations | Moderate; dependent on transcription and normalization |
| English lexical semantics | Zero validated confidence |
| Speculative English literary coherence | Creative criterion only, not evidential confidence |

## Falsification and review

A future decipherment claim should be rejected or revised if it cannot:

1. publish a stable glyph-to-value or token-to-value procedure;
2. apply that procedure to unseen folios without changing rules;
3. explain labels, running text, and section changes with the same core system;
4. produce independently checkable linguistic or cryptographic predictions;
5. outperform strong null models such as templated generation or local copying;
6. survive review by specialists in historical linguistics, cryptography, codicology, and the relevant medieval languages.

## Expansion gate

Version 0.1.0 freezes the pilot rule set. Expansion to the whole manuscript should happen only after documented review. Revisions must receive a new rule-set version and retain the old outputs so that changes cannot be hidden after seeing results.

