# Templates

This folder contains public, non-sensitive templates used by the extraction pipeline.

## `keyword_overrides.json`

Optional reviewed keyword extensions. Use this file when the default rules repeatedly miss a stable phrase in annual reports.

Supported sections:

| Section | Purpose |
| --- | --- |
| `field_groups.business` | Extra phrases for company overview, products, services and business descriptions. |
| `field_groups.segment` | Extra phrases for reportable segment or operating segment evidence. |
| `field_groups.revenue_table` | Extra phrases for revenue breakdown tables. |
| `field_groups.mda` | Extra phrases for MD&A / results of operations evidence. |
| `field_groups.revenue_recognition` | Extra phrases for revenue recognition evidence. |
| `table_keywords` | Extra words used when ranking relevant tables. |
| `theme_dictionary` | Extra theme keywords for AI, semiconductors, robotics, space, geopolitical risk, energy, minerals and biomedicine. |

Recommended rules:

- Keep additions short and reusable.
- Avoid company-specific names.
- Avoid broad risk words that appear in many unrelated filings.
- Keep project-specific notes outside the reusable keyword template.
- Review model-suggested keywords manually before adding them.

## `keyword_feedback_template.json`

A safe schema for model feedback when evidence is missing. The model should suggest keywords only; it should not edit code automatically.

Recommended workflow:

1. Run the rule-based extractor.
2. Review companies with weak or missing evidence.
3. Ask a model to produce feedback using this schema and the local evidence.
4. Manually review suggested keywords.
5. Add only stable, reusable terms to `keyword_overrides.json`.

Generated feedback may contain long annual report excerpts, so it is usually best kept with the run outputs instead of the reusable template folder.
