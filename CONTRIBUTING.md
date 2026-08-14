# Contributing

Contributions are welcome when they preserve the distinction between evidence and invention.

## Before opening a pull request

1. State whether the change affects source transcription, normalization, structural analysis, or creative reconstruction.
2. Cite the folio and locus identifiers affected.
3. Do not describe speculative English as plaintext, translation, or decipherment.
4. Add or update tests for machine-readable claim labels and row counts.
5. Run `python -m unittest discover -s tests -v`.
6. If changing a frozen rule, increment the rule-set version and preserve the previous release.

Claims of decipherment should include a stable procedure, unseen-page predictions, comparison against null models, and enough data for independent reproduction. Fluency or illustration fit alone is not validation.

