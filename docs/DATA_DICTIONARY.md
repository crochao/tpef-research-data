# Data dictionary

Multi-value fields use a semicolon followed by a space (`; `) unless otherwise stated. Empty values indicate unavailable or inapplicable information.

## LTND

Files: `data/TPEF_LTND_v1.0.xlsx` and `data/TPEF_LTND_v1.0.csv`

| Field | Description |
|---|---|
| `ID LTND` | Stable identifier assigned to the mechanically consolidated form. |
| `Lema representativo` | Display form selected from the observed source lemmas. |
| `Clave normalizada` | Comparison key produced by the normalization procedure. |
| `Formas originales observadas` | Distinct original forms, separated with ` || `. |
| `N.º formas originales` | Number of distinct original forms in the group. |
| `N.º entradas` | Number of corpus entries represented. |
| `N.º glosarios` | Number of source glossaries represented. |
| `Glosarios` | Source glossary codes. |
| `N.º países` | Number of countries represented. |
| `Países` | Three-letter country codes represented. |
| `Tipologías observadas` | Lexical-unit types recorded in the source data. |
| `Préstamos observados` | Loanword classifications recorded in the source data. |
| `Referencias de origen` | Source-row references that preserve traceability. |

## LTCAT groupings

File: `data/TPEF_LTCAT_agrupaciones_v1.0.csv`

| Field | Description |
|---|---|
| `ID LTCAT` | Identifier of the provisional terminological grouping. |
| `ID(s) v13`, `ID(s) v12`, `ID(s) v11` | Identifiers from earlier consolidation versions, retained for change tracking. |
| `Lema de agrupación provisional` | Provisional organizing lemma for the group. |
| `Variantes conservadas` | Equivalent or related denomination strings retained within the group. |
| `N.º denominaciones` | Number of denomination strings recorded for the group. |
| `N.º filas fuente` | Number of source inventory rows represented. |
| `Filas originales` | References to the source inventory rows. |
| `Tipología` | Lexical-unit type or combined types. |
| `Subdominio` | Primary thematic subdomain when assigned. |
| `Subdominio secundario` | Additional thematic subdomain when assigned. |
| `Préstamo` | Loanword status or classifications. |
| `Fuentes` | Glossary codes represented in the grouping. |
| `Países` | Country codes represented in the grouping. |
| `Validación original` | Validation information inherited from the working inventory. |
| `Estado / notas` | Editorial status and consolidation notes. |

## LTCAT denominations

File: `data/TPEF_LTCAT_denominaciones_v1.0.csv`

| Field | Description |
|---|---|
| `ID LTCAT` | Identifier of the parent grouping. |
| `Denominación` | Individual denomination string. |
| `Estatus` | Role within the group, such as organizing lemma or retained variant. |
| `Fuentes del grupo` | Glossary codes represented in the parent grouping. |
| `Países del grupo` | Country codes represented in the parent grouping. |

## LTCAT correspondences

File: `data/TPEF_LTCAT_correspondencias_v1.0.csv`

| Field | Description |
|---|---|
| `Fila fuente` | Row number in the source inventory used for consolidation. |
| `Lema original` | Lemma as recorded in that source inventory. |
| `ID LTCAT` | Identifier of the assigned provisional grouping. |
| `Lema consolidado` | Organizing lemma of the assigned grouping. |
| `ID(s) v13` | Identifier retained from consolidation version 13. |
| `Glosario` | Source glossary code. |
| `País` | Three-letter country code. |

## Glossary-source manifest

File: `data/TPEF_fuentes_glosarios_v1.0.csv`

| Field | Description |
|---|---|
| `Código` | Internal glossary code used throughout the study. |
| `País` | Country represented by the source. |
| `Institución` | Publishing institution. |
| `Título de glosario` | Title of the glossary or resource. |
| `Formato` | Web or document format from which the data were collected. |
| `Entradas` | Number of corpus entries contributed by the source. |
| `URL` | Source address documented during corpus compilation. |
