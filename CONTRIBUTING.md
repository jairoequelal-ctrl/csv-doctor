# Contributing

Please open an issue describing the import problem before a large change. Use synthetic fixtures: never upload private customer data.

Install with `python -m pip install -e ".[excel]"`, then run `python -m unittest discover -s tests -v`.

A repair must be explicit, preserve rejected values, account for quarantined records, and appear in the audit report. Include a regression fixture for parsing or transformation changes. Keep CSV-only runtime dependencies at zero and the interface local.

Helpful next contributions: additional explicit numeric conventions, schema validation, accessibility improvements, and localized diagnostic messages. Do not add silent type inference or external uploads.
