package execution;

import java.nio.file.Path;
import java.nio.file.Paths;

import com.test.qa.utils.excel.comparision.BulkFileComparator;
import com.test.qa.utils.excel.comparision.RuleConfigLoader;
import com.test.qa.utils.excel.comparision.ValidationProfile;

/** Runs the sample file-to-file comparison and writes both report formats. */
public class Test {

	/** Loads the sample profile and compares the configured sample files.
	 *
	 * @param args unused; sample file paths are configured in this class
	 *
	 * @throws Exception if input loading, comparison, or report writing fails */
	public static void main(String[] args) throws Exception {
		Path actualFile = Paths.get("input", "actual.csv");
		Path expectedFile = Paths.get("input", "expected.csv");
		Path reportFile = Paths.get("output", "comparison-report.xlsx");
		// Load your validation profile if required
		ValidationProfile profile = RuleConfigLoader.fromFile(Paths.get("config", "transaction-reconciliation-profile.json"));
		BulkFileComparator.Config config = new BulkFileComparator.Config();
		config.keyColumn = profile.getKeyColumn();
		// Transaction references are identifiers, not numbers; preserve their full text.
		config.keyType = BulkFileComparator.KeyType.STRING;
		config.chunkSize = 50_000;
		BulkFileComparator.compare(actualFile, expectedFile, reportFile, config, profile);
		System.out.println("Report generated: " + reportFile.toAbsolutePath());
		//#6uUdrFjd&gt
	}
}
