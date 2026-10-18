package com.test.qa.utils.excel.comparision;

/**
 * Dependency-free regression checks for comparison semantics. Run with: java
 * com.test.qa.utils.excel.comparision.ComparatorRegressionTest
 */
public final class ComparatorRegressionTest {
	private ComparatorRegressionTest() {
	}

	public static void main(String[] args) {
		whitespacePolicyIsRespected();
		decimalRoundingDoesNotHideDifferences();
		decimalToleranceIsAppliedToExactDifference();
		System.out.println("Comparator regression checks passed.");
	}

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

	private static void decimalRoundingDoesNotHideDifferences() {
		ColumnRule rule = new ColumnRule();
		rule.setType(ColumnRule.DataType.DECIMAL);
		rule.setScale(2);
		rule.setRoundingMode(java.math.RoundingMode.HALF_UP);
		rule.setTolerance("0");
		check(!RuleBasedValueComparator.compare("100.004", "100.003", rule).matches(),
				"Rounding must not hide a difference in exact source amounts");
	}

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

	private static void check(boolean condition, String message) {
		if (!condition)
			throw new AssertionError(message);
	}
}
