package com.test.qa.utils.excel.comparision;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Locale;

/** Applies a column's configured normalization and comparison semantics. */
public final class RuleBasedValueComparator {
	public record ComparisonResult(boolean matches, String normalizedActual, String normalizedExpected,
			String explanation) {
	}

	private RuleBasedValueComparator() {
	}

	public static ComparisonResult compare(String actual, String expected, ColumnRule rule) {
		if (rule == null)
			rule = new ColumnRule();
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

	private static String display(String value, ColumnRule rule) {
		if (value == null)
			return "<JAVA_NULL>";
		if (isExplicitNull(value, rule))
			return "<NULL_TOKEN>";
		return isBlank(value, rule) ? "<BLANK>" : clean(value, rule);
	}

	private static String clean(String value, ColumnRule rule) {
		if (!rule.isTrimWhitespace())
			return value;
		// If whitespace-only strings are configured as significant, don't trim them
		// into an empty string and accidentally make them equal to a blank value.
		if (!rule.isWhitespaceOnlyIsBlank() && !value.isEmpty() && value.trim().isEmpty())
			return value;
		return value.trim();
	}

	/** Parses numeric cells formatted by Excel with currency symbols or grouping separators. */
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

	private static String canonical(BigDecimal value) {
		BigDecimal normalized = value.stripTrailingZeros();
		return normalized.scale() < 0 ? normalized.setScale(0).toPlainString() : normalized.toPlainString();
	}

	private static String safeClean(String value, ColumnRule rule) {
		return value == null ? "<NULL>" : clean(value, rule);
	}
}
