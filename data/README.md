# Data dictionary and provenance

## Files

- `pilot_lines.csv` — one row per selected locus, suitable for spreadsheets.
- `pilot_lines.jsonl` — the same rows as newline-delimited JSON.
- `pilot_token_statistics.csv` — pilot and full-corpus token counts used by the structural glosses.
- `pilot_ruleset.json` — frozen folio selection, contextual descriptions, image mapping, and creative phrase inventories.
- `build_manifest.json` — source hash and build totals.

## `pilot_lines` fields

| Field | Meaning |
|---|---|
| `folio` | Manuscript surface identifier |
| `section_code` | Conventional section code from the source metadata |
| `section` | Human-readable conventional section name |
| `locator` | Stable ZL line/locus identifier |
| `position_code` | IVTFF position code |
| `locus_type` | Paragraph, label, circular text, radial text, or other encoded type |
| `raw_eva` | Source locus text preserved with IVTFF markup |
| `display_eva` | Human-readable display transformation |
| `normalized_tokens` | Token groups used for frequency calculations |
| `token_count` | Number of normalized token groups |
| `structural_gloss` | Nonsemantic form and position observations |
| `structural_confidence` | Confidence boundary for the analytical statement |
| `speculative_english` | Deterministic creative reconstruction |
| `speculative_status` | Mandatory warning that the line is not a translation |
| `page_context` | Illustration/layout context for the folio |
| `ruleset_version` | Generator and analysis rule-set identifier |

## Rights note

ZL3b-n is a modern scholarly transcription by René Zandbergen and collaborators. It is cited and hashed but not included as a complete source file. This repository does not relicense the source transcription. The pilot rows reproduce the minimum source strings needed for scholarly audit and pair them with new analysis and metadata.

The source file can be obtained from <https://www.voynich.nu/data/ZL3b-n.txt>.

