package execution;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import com.test.qa.utils.excel.comparision.BulkFileComparator;
import com.test.qa.utils.excel.comparision.JdbcQueryConfig;
import com.test.qa.utils.excel.comparision.RuleConfigLoader;
import com.test.qa.utils.excel.comparision.ValidationProfile;

/** Runs a database-versus-file comparison using the project's sample inputs and local settings. */
public class DatabaseTest {

	/** Loads settings, executes the configured query, and writes Excel and HTML reports.
	 *
	 * @param args optional actual file, SQL file, report file, and profile file paths in that order
	 *
	 * @throws Exception if settings, source files, database access, comparison, or report writing fails */
	public static void main(String[] args) throws Exception {
		Map<String, String> dotEnv = readDotEnv(Paths.get(".env"));
		Path actualFile = args.length > 0 ? Paths.get(args[0]) : Paths.get("input", "actual.csv");
		Path queryFile = args.length > 1 ? Paths.get(args[1]) : Paths.get("config", "expected-transactions.sql");
		Path reportFile = args.length > 2 ? Paths.get(args[2]) : Paths.get("output", "database-comparison-report.xlsx");
		Path profileFile = args.length > 3 ? Paths.get(args[3])
				: Paths.get("config", "transaction-reconciliation-profile.json");

		String query = Files.readString(queryFile, StandardCharsets.UTF_8);
		JdbcQueryConfig database = new JdbcQueryConfig(
				setting(JdbcQueryConfig.JDBC_URL_ENV, dotEnv),
				setting(JdbcQueryConfig.USERNAME_ENV, dotEnv),
				setting(JdbcQueryConfig.PASSWORD_ENV, dotEnv), query);
		ValidationProfile profile = RuleConfigLoader.fromFile(profileFile);

		BulkFileComparator.Config config = new BulkFileComparator.Config();
		config.keyColumn = profile.getKeyColumn();
		config.keyType = BulkFileComparator.KeyType.STRING;
		config.chunkSize = 50_000;

		BulkFileComparator.compare(actualFile, database, reportFile, config, profile);
		String reportName = reportFile.getFileName().toString();
		int extension = reportName.lastIndexOf('.');
		String baseName = extension > 0 ? reportName.substring(0, extension) : reportName;
		Path htmlReport = reportFile.resolveSibling(baseName + ".html");
		System.out.println("Excel report: " + reportFile.toAbsolutePath());
		System.out.println("HTML report: " + htmlReport.toAbsolutePath());
	}

	/** Resolves a setting from the process environment first, then the project .env file. */
	private static String setting(String name, Map<String, String> dotEnv) {
		String value = System.getenv(name);
		if (value == null || value.isBlank()) value = dotEnv.get(name);
		if (value == null || value.isBlank()) {
			throw new IllegalStateException("Required setting " + name
					+ " was not found in the process environment or project .env file");
		}
		return value;
	}

	/** Reads KEY=value entries from a .env file, allowing comments and matching quotes. */
	private static Map<String, String> readDotEnv(Path path) throws Exception {
		Map<String, String> values = new HashMap<>();
		if (!Files.exists(path)) return values;

		int lineNumber = 0;
		for (String line : Files.readAllLines(path, StandardCharsets.UTF_8)) {
			lineNumber++;
			String entry = line.trim();
			if (entry.isEmpty() || entry.startsWith("#")) continue;
			if (entry.startsWith("export ")) entry = entry.substring("export ".length()).trim();

			int separator = entry.indexOf('=');
			if (separator < 1) {
				throw new IllegalArgumentException("Invalid .env entry on line " + lineNumber
						+ "; expected KEY=value");
			}
			String name = entry.substring(0, separator).trim();
			String value = entry.substring(separator + 1).trim();
			if (value.length() >= 2) {
				char first = value.charAt(0);
				char last = value.charAt(value.length() - 1);
				if ((first == '"' && last == '"') || (first == '\'' && last == '\'')) {
					value = value.substring(1, value.length() - 1);
				}
			}
			values.put(name, value);
		}
		return values;
	}
}
