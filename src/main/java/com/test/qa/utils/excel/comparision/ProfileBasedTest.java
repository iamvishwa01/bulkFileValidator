package com.test.qa.utils.excel.comparision;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;

/** Command-line entry point for profile-based file or database comparisons. */
public final class ProfileBasedTest {
	/** Prevents construction of this command-line utility class. */
	private ProfileBasedTest() {
	}

	/** Parses the selected mode and runs a comparison.
	 *
	 * @param args file mode: actual, expected, report, optional profile; DB mode: --db, actual, report, SQL, optional profile
	 *
	 * @throws Exception if configuration loading, comparison, or report generation fails */
	public static void main(String[] args) throws Exception {
		if (args.length > 0 && "--db".equals(args[0])) {
			if (args.length < 4 || args.length > 5) {
				usage();
				System.exit(2);
			}
			Path actual = Paths.get(args[1]);
			Path report = Paths.get(args[2]);
			String query = Files.readString(Paths.get(args[3]), StandardCharsets.UTF_8);
			ValidationProfile profile = loadProfile(args.length == 5 ? args[4] : null);
			BulkFileComparator.compare(actual, JdbcQueryConfig.fromEnvironment(query), report, config(), profile);
			return;
		}

		if (args.length < 3 || args.length > 4) {
			usage();
			System.exit(2);
		}
		Path actual = Paths.get(args[0]);
		Path expected = Paths.get(args[1]);
		Path report = Paths.get(args[2]);
		ValidationProfile profile = loadProfile(args.length == 4 ? args[3] : null);
		BulkFileComparator.compare(actual, expected, report, config(), profile);
	}

	/** Creates the common comparison settings used by the command-line modes. */
	private static BulkFileComparator.Config config() {
		BulkFileComparator.Config config = new BulkFileComparator.Config();
		// Transaction references are identifiers; preserve leading zeros by default.
		config.keyType = BulkFileComparator.KeyType.STRING;
		config.chunkSize = 50_000;
		config.caseInsensitiveHeaders = true;
		config.trimHeaders = true;
		config.strictHeaderOrder = false;
		return config;
	}

	/** Loads a user-specified profile or the bundled sample profile. */
	private static ValidationProfile loadProfile(String profilePath) throws Exception {
		return profilePath == null ? RuleConfigLoader.fromResource("profiles/sample-validation-profile.json")
				: RuleConfigLoader.fromFile(Paths.get(profilePath));
	}

	/** Prints supported command-line forms and database environment variable names. */
	private static void usage() {
		System.err.println("File mode: ProfileBasedTest <actual.csv|xlsx> <expected.csv|xlsx> <report.xlsx> [profile.json]");
		System.err.println("Database mode: ProfileBasedTest --db <actual.csv|xlsx> <report.xlsx> <query.sql> [profile.json]");
		System.err.println("Database mode reads " + JdbcQueryConfig.JDBC_URL_ENV + ", "
				+ JdbcQueryConfig.USERNAME_ENV + ", and " + JdbcQueryConfig.PASSWORD_ENV + " from the environment.");
		System.err.println("An HTML report with the same base name is also created.");
	}
}
