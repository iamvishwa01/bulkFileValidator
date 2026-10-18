package com.test.qa.utils.excel.comparision;

import java.nio.file.Path;
import java.nio.file.Paths;

/** Example entry point showing profile-based and per-column validation. */
public final class ProfileBasedTest {
	private ProfileBasedTest() {
	}

	public static void main(String[] args) throws Exception {
		if (args.length < 3 || args.length > 4) {
			System.err.println(
					"Usage: ProfileBasedTest <actual.csv|xlsx> <expected.csv|xlsx> <report.xlsx> [profile.json]");
			System.exit(2);
		}

		Path actual = Paths.get(args[0]);
		Path expected = Paths.get(args[1]);
		Path report = Paths.get(args[2]);

		BulkFileComparator.Config config = new BulkFileComparator.Config();
		// Transaction references are identifiers; preserve leading zeros by default.
		config.keyType = BulkFileComparator.KeyType.STRING;
		config.chunkSize = 50_000;
		config.caseInsensitiveHeaders = true;
		config.trimHeaders = true;
		config.strictHeaderOrder = false;

		ValidationProfile profile = args.length == 4 ? RuleConfigLoader.fromFile(Paths.get(args[3]))
				: RuleConfigLoader.fromResource("profiles/sample-validation-profile.json");

		BulkFileComparator.compare(actual, expected, report, config, profile);
	}
}
