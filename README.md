# TPEF research data and prototype

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.22047082.svg)](https://doi.org/10.5281/zenodo.22047082)

This repository accompanies the doctoral dissertation on the terminographic representation of financial education in the Spanish-speaking world. It contains derived datasets, reproducible analysis scripts, and selected SKOS/RDF examples for the *Tesauro panhispánico de educación financiera* (TPEF).

## Release summary

- Version: 1.0.0
- Release date: 8 October 2026
- Corpus scope: 18,933 entries from 57 institutional glossaries in 18 Spanish-speaking countries
- LTND: 10,846 normalized and mechanically deduplicated lemma forms
- LTCAT: 9,076 provisional terminological groupings containing 11,516 denomination strings

The LTCAT is a consolidated terminological inventory produced through automated processing and manual analysis. It remains provisional because some equivalences, polysemous senses, preferred labels, and conceptual relations require expert validation. The SKOS files are representative samples rather than a complete operational thesaurus.

## Repository contents

### `data/`

- `TPEF_LTND_v1.0.xlsx` and `.csv`: Total Normalized and Deduplicated Lemma Inventory (*Lemario total normalizado y deduplicado*, LTND).
- `TPEF_LTCAT_v1.0.xlsx`: Total Lemma Inventory Consolidated through Terminological Analysis (*Lemario total consolidado por análisis terminológico*, LTCAT).
- `TPEF_LTCAT_agrupaciones_v1.0.csv`: one record per provisional LTCAT grouping.
- `TPEF_LTCAT_denominaciones_v1.0.csv`: long-form table of denominations and their status within each grouping.
- `TPEF_LTCAT_correspondencias_v1.0.csv`: traceability table linking source rows to consolidated groupings.
- `TPEF_fuentes_glosarios_v1.0.csv`: manifest of the 57 institutional glossaries.

### `scripts/`

- `tpef_generar_ltnd.py`: rebuilds the LTND from the authorized source workbook.
- `tpef_jaccard.py`: calculates lexical overlap between glossaries and countries.
- `tpef_denominaciones.py`: extracts candidate equivalent denominations from lemmas and definitions.
- `tpef_geosinonimia_sketchengine.py`: queries esTenTen23 through the Sketch Engine API.
- `tpef_clasificador_subdominios.py`: assigns entries to eight thematic subdomains by weighted rules.
- `tpef_clustering_conceptos.py`: generates candidate conceptual groups using TF-IDF and cosine similarity.
- `tpef_jerarquias.py`: identifies candidate hierarchical structures from multiword lemmas.

### `rdf/`

The folder contains four validated Turtle files and matching RDF/XML serializations. The two principal samples contain 106 triples for the current-account network and 181 triples for the credit network. The mortgage and brand-metonymy files provide supplementary modeling examples.

### `docs/`

- `DATA_DICTIONARY.md`: field-level documentation.
- `METHODOLOGY.md`: processing levels, scope, and reproducibility notes.
- `ZENODO_METADATA_DRAFT.md`: metadata ready to adapt when creating the Zenodo record.
- `FILE_MANIFEST.md`: concise inventory of the release.

## Reproducing the LTND

The raw source workbook is excluded from this public package because the glossary definitions originate from third-party institutional resources whose reuse conditions vary. Authorized users can place `Glosarios_2026.xlsx` in the project root and run:

```bash
python scripts/tpef_generar_ltnd.py Glosarios_2026.xlsx --output-dir data/rebuilt_ltnd
```

The expected result is 10,846 LTND records from 18,933 corpus entries. The program preserves traceability to the source glossary and country.

## Software environment

Python 3.10 or later is recommended. Install the dependencies with:

```bash
python -m pip install -r requirements.txt
```

The Sketch Engine script additionally requires an authorized account and the environment variables `SKETCHENGINE_USERNAME` and `SKETCHENGINE_API_KEY`.

## Scope and reuse cautions

- The source manifest documents the glossaries; it does not redistribute their definitions.
- The thematic classifier and clustering procedures generate candidates for terminological review.
- The namespace `https://w3id.org/tpef/` is used in the prototype files. Long-term persistence requires registration and continued maintenance.
- The release follows a FAIR-oriented architecture. Full FAIR compliance depends on persistent publication, licensing, documentation, and version maintenance.

## Licensing

- The Python scripts are licensed under the MIT License.
- The derived datasets, RDF files, and documentation are licensed under the Creative Commons Attribution 4.0 International License (CC BY 4.0).
- Copyright and database rights in the institutional source resources remain with their respective holders. Their glossary definitions are not redistributed in this release.

See `LICENSE`, `LICENSE-CODE`, and `LICENSE-DATA` for details.

## Publication records

- GitHub: <https://github.com/crochao/tpef-research-data/>
- Zenodo: <https://doi.org/10.5281/zenodo.22047082>

## Remaining preservation checks

1. Confirm the status of the `w3id.org/tpef/` namespace.
2. Review institutional reuse restrictions and any dissertation embargo.

## Citation

Use the citation displayed on the [Zenodo record](https://doi.org/10.5281/zenodo.22047082). A machine-readable citation is supplied in `CITATION.cff`.

## Author

Carlos O. Rocha Ochoa — <https://orcid.org/0000-0002-2475-6238>
