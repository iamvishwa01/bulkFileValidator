# Bulk File Validator

Java 17 bulk CSV/XLSX comparison engine with bounded-memory external sorting and optional JSON-based per-column validation rules.

## What changed

- Kept the existing `compare(actual, expected, report, config)` API working.
- Added `compare(actual, expected, report, config, profile)` for per-column rules.
- Profile key column is applied to a copied `Config`, so the caller's configuration is not mutated (important for parallel comparisons).
- For matched rows, rules control text case sensitivity, whitespace trimming, decimal scale/rounding, tolerance, explicit null tokens and null-vs-blank behavior.
- Mismatch report values show the normalized values and the rule comparator's explanation.
- Added a sample JSON profile, runnable example and unit tests for comparator semantics.

## Run

Use Maven with Java 17:

```bash
mvn test
mvn package
```

Then run with the sample profile on the classpath:

```bash
java -cp target/classes:target/dependency/* com.test.qa.utils.excel.comparision.ProfileBasedTest actual.csv expected.xlsx output/report.xlsx
```

For a custom JSON profile:

```bash
java -cp target/classes:target/dependency/* com.test.qa.utils.excel.comparision.ProfileBasedTest actual.csv expected.xlsx output/report.xlsx C:/config/card-report.json
```

The runtime classpath must include dependencies. On Windows, use `;` rather than `:` between classpath entries. The simplest local workflow is to run `ProfileBasedTest` from your IDE with the Maven project imported.

## Profile example

See `config/sample-validation-profile.json` or the packaged `src/main/resources/profiles/sample-validation-profile.json`.
The sample profile matches the employee-style columns in `input/actual.xlsx` and `input/expected.xlsx`.
Column names are matched case-insensitively by `ValidationProfile.ruleFor`; columns not listed use default `ColumnRule` behavior when a profile is supplied.

## Important behavior

- Decimal scale and rounding are validated, while tolerance is applied to the exact source values. With `HALF_UP`, scale 2 and tolerance `0.01`, `10.00` and `10.01` match; `10.00` and `10.02` do not. Rounding cannot hide a difference outside the configured tolerance.
- `roundingMode: UNNECESSARY` rejects values that cannot be represented at the configured scale without rounding.
- `nullEqualsBlank` defaults to `false`; enable it only when the report specification says explicit null tokens may equal blanks.
- The current comparator compares the first worksheet in an XLSX file, matching the existing implementation.

## Simpler comparator

For a smaller, easier-to-follow implementation, see `com.test.qa.utils.excel.simple.SimpleFileComparator`.
It reads CSV or the first worksheet of XLSX files, compares rows by a text key, and writes a CSV mismatch report.
Text keys preserve leading zeros (`001234` is different from `1234`). Duplicate keys are reported without pairing.
Pass decimal tolerances by column name when exact numeric comparison is needed; without a tolerance, values are
compared as trimmed text. This version loads both files into memory, so use the original comparator for large files.

Example from Java code:

```java
SimpleFileComparator.compare(
    Path.of("input/actual.csv"),
    Path.of("input/expected.xlsx"),
    Path.of("output/mismatches.csv"),
    "ID",
    Map.of("Amount", new BigDecimal("0.01"))
);
```
