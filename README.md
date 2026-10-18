# Bulk File Validator

Bulk File Validator compares two CSV or Excel files and reports differences by row key. It is designed for large files: input rows are processed in bounded chunks, sorted into temporary runs, then merged for comparison. A JSON validation profile can define different comparison rules for individual columns.

## Capabilities

- Compare `.csv` files and `.xlsx` workbooks. For workbooks, the first worksheet is used.
- Match records using a configured key column. Keys can be sorted as text or as numbers.
- Compare headers without regard to case and surrounding whitespace by default. Column order does not need to match unless strict header ordering is enabled.
- Compare values using global settings or per-column rules from a JSON profile.
- Identify missing and extra rows, duplicate keys, mismatched values, header problems, malformed records, and blank keys.
- Write a two-sheet XLSX report and a self-contained HTML report with the same base name.
- Redact values in sensitive columns and PAN-like digit sequences in both reports.

## Requirements

- Java 17 or later
- Maven 3.8 or later

The project uses Apache Commons CSV to read CSV files, Apache POI to stream XLSX files and write reports, and Jackson to load validation profiles.

## Build

From the project directory, run:

```bash
mvn clean package
```

To copy runtime dependencies for direct command-line execution:

```bash
mvn dependency:copy-dependencies
```

## Run a comparison

The `ProfileBasedTest` entry point accepts the actual file, expected file, XLSX report path, and an optional JSON profile. Pass a profile path to apply per-column rules. The no-profile form uses the built-in classpath resource `profiles/sample-validation-profile.json` when that resource is present on the runtime classpath. The repository's editable profile files are in `config/`, so pass one of those paths for the examples below.

```text
ProfileBasedTest <actual.csv|actual.xlsx> <expected.csv|expected.xlsx> <report.xlsx> [profile.json]
```

For example, on Windows:

```powershell
java -cp "target/classes;target/dependency/*" com.test.qa.utils.excel.comparision.ProfileBasedTest input/actual.csv input/expected.csv output/comparison-report.xlsx config/transaction-reconciliation-profile.json
```

On Linux or macOS, use a colon between classpath entries:

```bash
java -cp "target/classes:target/dependency/*" com.test.qa.utils.excel.comparision.ProfileBasedTest input/actual.csv input/expected.csv output/comparison-report.xlsx config/transaction-reconciliation-profile.json
```

The project also includes `execution.Test` under `src/test/java`. It is a runnable `main` program configured for the sample transaction files and transaction profile; it is not a JUnit test. In an IDE, run `execution.Test` as a Java application. When run from the project directory, it compares `input/actual.csv` with `input/expected.csv` and writes reports under `output/`.

For a different profile, pass its path as the fourth argument to `ProfileBasedTest` or change the profile path in `execution.Test`.

## Reports

If the requested report is `output/comparison-report.xlsx`, the comparison also writes `output/comparison-report.html`.

The XLSX workbook contains:

- **Summary** — source file paths, row counts, mismatch totals, execution status, and the final PASS/FAIL result.
- **Mismatches** — one row per reported issue, with type, key, column, actual value, expected value, and details.

The HTML report presents the same summary and mismatch information in a browser-friendly layout. It is a standalone file and requires no external assets or network access.

The comparison passes only when there are no header mismatches, value mismatches, missing rows, extra rows, duplicate keys, or malformed rows. A schema problem such as a missing key column stops evaluation and is reported as a schema failure. Processing errors are recorded in a report when possible and are also raised to the caller.

## Row matching and duplicate keys

Rows are sorted by the key column before comparison, so input files do not need to be in the same order. String keys use lexical ordering and retain leading zeros (`00123` differs from `123`). Numeric keys use decimal numeric ordering (`20` sorts before `100`). Choose string keys for identifiers, account references, and codes whose formatting matters.

Each key is expected to occur once in each file. Duplicate key groups are reported and are not paired arbitrarily. When the same key occurs more than once in either file, the report marks the match as ambiguous. A blank key is invalid and its row is skipped. Rows with a different field count than the header are also reported and skipped.

## Validation profiles

A profile is a JSON object with a required `profileName`, a required `keyColumn`, and a `columns` object containing rules keyed by column name. Column names are matched without regard to case. Columns omitted from the profile use the default string rule.

The repository includes:

- `config/sample-validation-profile.json` — example employee-style rules.
- `config/transaction-reconciliation-profile.json` — rules for transaction reconciliation, including amount tolerances.

Profile fields:

| Field | Meaning |
| --- | --- |
| `type` | `STRING` or `DECIMAL`. Defaults to `STRING`. |
| `trimWhitespace` | Trim surrounding whitespace before comparing. Defaults to `true`. |
| `caseSensitive` | For string values, compare letter case. Defaults to `true`. |
| `nullEqualsBlank` | Allow an explicit null value to equal a blank. Defaults to `false`. |
| `whitespaceOnlyIsBlank` | Treat values containing only whitespace as blank. Defaults to `true`. |
| `nullTokens` | Text tokens, such as `NULL` or `N/A`, to treat as explicit nulls. |
| `scale` | For decimals, validate representability at this scale. Range: 0–12. |
| `roundingMode` | Java rounding mode used for the decimal scale check, such as `HALF_UP` or `UNNECESSARY`. |
| `tolerance` | Maximum absolute difference allowed for decimals. Must be non-negative. Defaults to `0`. |

Example:

```json
{
  "profileName": "transaction-reconciliation",
  "keyColumn": "Payment / Auth ID",
  "columns": {
    "Status": {
      "type": "STRING",
      "trimWhitespace": true,
      "caseSensitive": false,
      "nullEqualsBlank": false,
      "whitespaceOnlyIsBlank": true,
      "nullTokens": ["NULL", "N/A"]
    },
    "Gross Amount": {
      "type": "DECIMAL",
      "scale": 2,
      "roundingMode": "HALF_UP",
      "tolerance": "0.01"
    }
  }
}
```

For decimal rules, scale and rounding determine whether a value can be represented according to the profile. The comparator still checks the exact source values against the tolerance, so rounding alone cannot conceal a difference. For example, with scale 2, `HALF_UP`, and tolerance `0.01`, `10.00` and `10.01` match, while `10.00` and `10.02` do not. `UNNECESSARY` reports a mismatch when a value needs rounding to fit the configured scale.

## Java API

The primary API accepts `Path` values, a `BulkFileComparator.Config`, and optionally a `ValidationProfile`:

```java
Path actual = Path.of("input/actual.csv");
Path expected = Path.of("input/expected.csv");
Path xlsxReport = Path.of("output/comparison-report.xlsx");

BulkFileComparator.Config config = new BulkFileComparator.Config();
config.keyColumn = "Payment / Auth ID";
config.keyType = BulkFileComparator.KeyType.STRING;
config.chunkSize = 50_000;

ValidationProfile profile = RuleConfigLoader.fromFile(
    Path.of("config/transaction-reconciliation-profile.json"));

BulkFileComparator.compare(actual, expected, xlsxReport, config, profile);
```

For global comparison behavior without a profile, call `BulkFileComparator.compare(actual, expected, report, config)`. The supplied configuration is copied before the comparison runs, so profile settings do not mutate the caller's `Config` instance.

## Operational notes

- `chunkSize` controls how many rows are sorted in memory at a time. Lower values reduce memory use but can create more temporary sort files.
- Temporary sort data is created in the system temporary directory and removed when a comparison finishes.
- CSV files use a comma delimiter by default. The first row is treated as the header.
- XLSX values are read as formatted cell text, and only the first worksheet is compared.
- Header matching trims surrounding whitespace and ignores case by default. Header order is ignored unless strict ordering is enabled.
- Sensitive column names, including names containing `PAN`, `card number`, `CVV`, `PIN`, or `security code`, are redacted in reports. PAN-like digit sequences are also redacted as a safeguard.
