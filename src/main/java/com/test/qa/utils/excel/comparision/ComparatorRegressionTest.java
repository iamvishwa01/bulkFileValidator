package com.test.qa.utils.excel.comparision;

/** Standalone regression checks for the value-comparison policies. */
public final class ComparatorRegressionTest {
	/** Prevents construction of this utility class. */
	private ComparatorRegressionTest() {
	}

	/** Runs each comparison-policy check and prints a success message. */
	/**
	 * @param args unused */
	public static void main(String[] args) {
		whitespacePolicyIsRespected();
		decimalRoundingDoesNotHideDifferences();
		decimalToleranceIsAppliedToExactDifference();
		System.out.println("Comparator regression checks passed.");
	}

	/** Checks that whitespace-only values follow the configured blank policy. */
	private static void whitespacePolicyIsRespected() {
		ColumnRule rule = new ColumnRule();
		rule.setType(ColumnRule.DataType.STRING);
		rule.setTrimWhitespace(true);
		rule.setWhitespaceOnlyIsBlank(false);
		check(!RuleBasedValueComparator.compare(" ", "", rule).matches(),
				"Whitespace-only value must not become blank when policy says significant");

		rule.setWhitespaceOnlyIsBlank(true);
		check(RuleBasedValueComparator.compare(" ", "", rule).matches(),
				"Whitespace-only value should be blank when policy says blank");
	}

	/** Checks that decimal rounding does not conceal differences in source values. */
	private static void decimalRoundingDoesNotHideDifferences() {
		ColumnRule rule = new ColumnRule();
		rule.setType(ColumnRule.DataType.DECIMAL);
		rule.setScale(2);
		rule.setRoundingMode(java.math.RoundingMode.HALF_UP);
		rule.setTolerance("0");
		check(!RuleBasedValueComparator.compare("100.004", "100.003", rule).matches(),
				"Rounding must not hide a difference in exact source amounts");
	}

	/** Checks that tolerance is compared against the exact decimal difference. */
	private static void decimalToleranceIsAppliedToExactDifference() {
		ColumnRule rule = new ColumnRule();
		rule.setType(ColumnRule.DataType.DECIMAL);
		rule.setScale(2);
		rule.setRoundingMode(java.math.RoundingMode.HALF_UP);
		rule.setTolerance("0.01");
		check(RuleBasedValueComparator.compare("100.004", "100.003", rule).matches(),
				"Configured tolerance should be applied to the exact decimal difference");
		check(!RuleBasedValueComparator.compare("100.02", "100.00", rule).matches(),
				"Difference above tolerance must fail");
	}

	/** Fails a regression check when its condition is false. */
	private static void check(boolean condition, String message) {
		if (!condition)
			throw new AssertionError(message);
	}
}
