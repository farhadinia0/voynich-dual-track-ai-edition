# Reproducibility

## Environment

- Python 3.11+
- `python-docx==1.2.0`
- `Pillow==12.3.0`

Install the pinned packages with:

```bash
python -m pip install -r requirements.txt
```

## Acquire the transcription

```bash
python src/fetch_zl.py
```

The fetcher downloads the versioned ZL3b-n source and refuses to save it unless its SHA-256 equals the value recorded in the source code and manifest. If the upstream file changes, archive the new input under a new project version; do not silently update the expected hash.

## Rebuild

From the repository root:

```bash
python src/build_pilot.py \
  --zl sources/ZL3b-n.txt \
  --image-dir assets/folios \
  --repo .
```

The build regenerates:

- `data/pilot_lines.csv`
- `data/pilot_lines.jsonl`
- `data/pilot_token_statistics.csv`
- `data/pilot_ruleset.json`
- `data/build_manifest.json`
- both Word editions in `editions/`

## Validate

```bash
python -m unittest discover -s tests -v
```

The tests check row counts, folio coverage, source hash metadata, claim labels, nonempty provenance fields, and required disclaimers in both Word files.

## Determinism

The data tables and speculative English lines are deterministic for a fixed source file, script, rule-set version, and phrase inventory. DOCX package bytes may differ across environments because office-file metadata and ZIP ordering can vary; validate content and structure, not only the final DOCX hash.

## Provenance chain

```text
Yale manuscript image / cited facsimile
             +
ZL3b-n IVTFF source (pinned SHA-256)
             |
             v
      parse and normalize
             |
       +-----+-------------------+
       |                         |
       v                         v
structural observations   frozen creative generator
       |                         |
       v                         v
 analytical edition       speculative edition
```

