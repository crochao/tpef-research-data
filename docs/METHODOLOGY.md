# Methodology and reproducibility notes

## Processing levels

The release distinguishes three analytical levels:

1. **Corpus entries**: 18,933 independently preserved lexicographic records.
2. **LTND**: 10,846 forms obtained through automated normalization and mechanical deduplication.
3. **LTCAT**: 9,076 provisional groupings produced through terminological consolidation, containing 11,516 denomination strings.

The LTND supports macrostructural counts and lexical-overlap analysis. The LTCAT resolves residual graphic and morphosyntactic redundancy, decomposes lemmas containing several denominations, preserves variants, and separates recognized collisions between abbreviated forms. Its group count should not be treated as a definitive concept count.

## Computational procedures

- **Lexical overlap** is calculated from normalized lemma sets with the Jaccard index.
- **Equivalent-denomination extraction** uses patterns found in lemmas and definitions.
- **Geosynonymy assessment** contrasts candidates in the national web domains of esTenTen23 through the Sketch Engine API.
- **Thematic classification** combines weighted rules, contextual evidence, and manual review across eight subdomains.
- **Conceptual estimation** groups selected definitions through TF-IDF vectors and cosine similarity.
- **Hierarchical candidates** are obtained from recurring initial components in multiword lemmas.

Each computational result is interpreted as evidence for review. Formal similarity, shared wording, and thematic proximity do not by themselves establish conceptual identity.

## Source-data availability

The source workbook contains third-party glossary definitions. It is omitted from the public release pending verification of the reuse conditions for every source. The source manifest supplies institutional provenance and URLs. The LTND generator is included so that authorized users can reproduce the derived inventory.

## Expected reproducibility checks

- Input corpus: 18,933 entries.
- Glossaries: 57.
- Countries: 18.
- LTND output: 10,846 records.
- Provisional LTCAT output: 9,076 groupings and 11,516 denomination strings.
- Principal SKOS samples: 106 triples for current account and 181 triples for credit.

## Versioning policy

Published versions should be immutable. Corrections should receive a new semantic version, an updated changelog, regenerated checksums, and a new Zenodo version linked to the preceding record.
