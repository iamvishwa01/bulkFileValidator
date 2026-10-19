package com.test.qa.utils.excel.comparision;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Locale;

/** Normalizes and compares two cell values according to a column rule. */
public final class RuleBasedValueComparator {
	/** Contains the comparison outcome, normalized values, and a human-readable explanation. */
	public record ComparisonResult(boolean matches, String normalizedActual, String normalizedExpected,
			String explanation) {
	}

	/** Prevents construction of this static utility class. */
	private RuleBasedValueComparator() {
	}

	/** Compares two values using blank, null, string, or decimal policies from the rule.
	 *
	 * @param actual value from the actual source, possibly null
	 *
	 * @param expected value from the expected source, possibly null
	 *
	 * @param rule column policy; a default rule is used when null
	 *
	 * @return comparison result and normalized display values */
	public static ComparisonResult compare(String actual, String expected, ColumnRule rule) {
		if (rule == null)
			rule = new ColumnRule();
		if (rule.isIgnoreExpectedBlank() && (isBlank(expected, rule) || isExplicitNull(expected, rule))) {
			return new ComparisonResult(true, display(actual, rule), display(expected, rule),
					"Expected value is blank; comparison skipped by the column rule");
		}
		boolean actualExplicitNull = isExplicitNull(actual, rule);
		boolean expectedExplicitNull = isExplicitNull(expected, rule);
		if (actualExplicitNull || expectedExplicitNull) {
			boolean otherSideIsBlank = actualExplicitNull ? isBlank(expected, rule) : isBlank(actual, rule);
			boolean match = actualExplicitNull && expectedExplicitNull
					|| (rule.isNullEqualsBlank() && otherSideIsBlank);
			return new ComparisonResult(match, display(actual, rule), display(expected, rule),
					match ? "Values match under the configured null/blank policy"
							: "Null/blank mismatch under the configured policy");
		}

		boolean actualBlank = isBlank(actual, rule);
		boolean expectedBlank = isBlank(expected, rule);
		if (actualBlank || expectedBlank) {
			boolean matches = actualBlank && expectedBlank;
			return new ComparisonResult(matches, display(actual, rule), display(expected, rule),
					matches ? "Both values are blank under the configured policy"
							: "Blank/non-blank mismatch under the configured policy");
		}

		if (rule.getType() == ColumnRule.DataType.DECIMAL) {
			try {
				BigDecimal rawActual = parseDecimal(clean(actual, rule));
				BigDecimal rawExpected = parseDecimal(clean(expected, rule));
				// Apply scale/rounding as a configuration validity check, but compare
				// the original exact values so rounding cannot hide a financial delta.
				RoundingMode mode = rule.getRoundingMode();
				rawActual.setScale(rule.getScale(), mode);
				rawExpected.setScale(rule.getScale(), mode);
				BigDecimal delta = rawActual.subtract(rawExpected).abs();
				BigDecimal tolerance = new BigDecimal(rule.getTolerance());
				boolean matches = delta.compareTo(tolerance) <= 0;
				return new ComparisonResult(matches, canonical(rawActual), canonical(rawExpected),
						matches ? "Exact decimal difference " + delta + " is within tolerance " + tolerance
								: "Exact decimal difference " + delta + " exceeds tolerance " + tolerance);
			} catch (NumberFormatException | ArithmeticException ex) {
				return new ComparisonResult(false, safeClean(actual, rule), safeClean(expected, rule),
						"Invalid decimal or configured rounding cannot be applied: " + ex.getMessage());
			}
		}

		String a = clean(actual, rule);
		String e = clean(expected, rule);
		if (!rule.isCaseSensitive()) {
			a = a.toLowerCase(Locale.ROOT);
			e = e.toLowerCase(Locale.ROOT);
		}
		boolean matches = a.equals(e);
		return new ComparisonResult(matches, a, e, matches ? "Text values match after configured normalization"
				: "Text values differ after configured normalization");
	}

	/** Tests whether a value is Java null or matches a configured null token. */
	private static boolean isExplicitNull(String value, ColumnRule rule) {
		if (value == null)
			return true;
		String candidate = rule.isTrimWhitespace() ? value.trim() : value;
		for (String token : rule.getNullTokens()) {
			if (token != null && candidate.equalsIgnoreCase(token))
				return true;
		}
		return false;
	}

	/** Tests whether a value is empty or configured whitespace-only text. */
	private static boolean isBlank(String value, ColumnRule rule) {
		if (value == null)
			return false;
		if (value.isEmpty())
			return true;
		boolean whitespaceOnly = value.trim().isEmpty();
		if (whitespaceOnly)
			return rule.isWhitespaceOnlyIsBlank();
		return false;
	}

	/** Produces a readable marker for null and blank values for mismatch reports. */
	private static String display(String value, ColumnRule rule) {
		if (value == null)
			return "<JAVA_NULL>";
		if (isExplicitNull(value, rule))
			return "<NULL_TOKEN>";
		return isBlank(value, rule) ? "<BLANK>" : clean(value, rule);
	}

	/** Applies the configured whitespace normalization to a non-null value. */
	private static String clean(String value, ColumnRule rule) {
		if (!rule.isTrimWhitespace())
			return value;
		// If whitespace-only strings are configured as significant, don't trim them
		// into an empty string and accidentally make them equal to a blank value.
		if (!rule.isWhitespaceOnlyIsBlank() && !value.isEmpty() && value.trim().isEmpty())
			return value;
		return value.trim();
	}

	/** Parses numeric text, supporting currency symbols, grouping marks, and accounting negatives. */
	private static BigDecimal parseDecimal(String value) {
		String normalized = value.replaceAll("[\\p{Sc}\\s\\u00A0\\u202F]", "");
		boolean accountingNegative = normalized.startsWith("(") && normalized.endsWith(")");
		if (accountingNegative) {
			normalized = normalized.substring(1, normalized.length() - 1);
		}

		int lastComma = normalized.lastIndexOf(',');
		int lastDot = normalized.lastIndexOf('.');
		if (lastComma >= 0 && lastDot >= 0) {
			// The rightmost separator is the decimal mark; the other is a thousands mark.
			if (lastComma > lastDot) {
				normalized = normalized.replace(".", "").replace(',', '.');
			} else {
				normalized = normalized.replace(",", "");
			}
		} else if (lastComma >= 0) {
			// A three-digit final group is a thousands separator; otherwise treat comma as decimal.
			if (normalized.matches(".*,[0-9]{3}")) {
				normalized = normalized.replace(",", "");
			} else {
				normalized = normalized.replace(',', '.');
			}
		}

		BigDecimal result = new BigDecimal(normalized);
		return accountingNegative ? result.negate() : result;
	}

	/** Converts a decimal to a stable plain-text representation without trailing zeroes. */
	private static String canonical(BigDecimal value) {
		BigDecimal normalized = value.stripTrailingZeros();
		return normalized.scale() < 0 ? normalized.setScale(0).toPlainString() : normalized.toPlainString();
	}

	/** Returns a normalized value while preserving a safe marker for null. */
	private static String safeClean(String value, ColumnRule rule) {
		return value == null ? "<NULL>" : clean(value, rule);
	}
}
