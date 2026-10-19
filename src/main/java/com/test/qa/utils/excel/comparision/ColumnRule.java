package com.test.qa.utils.excel.comparision;

import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;

/** Defines normalization, null handling, and value comparison for one column. */
public class ColumnRule {
	/** Selects text equality or exact decimal comparison for a column. */
	public enum DataType {
		STRING, DECIMAL
	}

	private DataType type = DataType.STRING;
	private boolean trimWhitespace = true;
	private boolean caseSensitive = true;
	private boolean nullEqualsBlank = false;
	private boolean whitespaceOnlyIsBlank = true;
	private boolean ignoreExpectedBlank = false;
	private int scale = 2;
	private RoundingMode roundingMode = RoundingMode.UNNECESSARY;
	private String tolerance = "0";
	private List<String> nullTokens = new ArrayList<>();

	/** Returns the configured value type. */
	public DataType getType() {
		return type;
	}

	/** Sets the value type, defaulting to {@link DataType#STRING} when null.
	 *
	 * @param type comparison type to use */
	public void setType(DataType type) {
		this.type = type == null ? DataType.STRING : type;
	}

	/** Returns whether surrounding whitespace is removed before comparison. */
	public boolean isTrimWhitespace() {
		return trimWhitespace;
	}

	/** Controls removal of surrounding whitespace before comparison.
	 *
	 * @param trimWhitespace true to trim values */
	public void setTrimWhitespace(boolean trimWhitespace) {
		this.trimWhitespace = trimWhitespace;
	}

	/** Returns whether string comparison is case-sensitive. */
	public boolean isCaseSensitive() {
		return caseSensitive;
	}

	/** Sets whether string comparison distinguishes letter case.
	 *
	 * @param caseSensitive true to distinguish letter case */
	public void setCaseSensitive(boolean caseSensitive) {
		this.caseSensitive = caseSensitive;
	}

	/** Returns whether explicit null values may equal blank values. */
	public boolean isNullEqualsBlank() {
		return nullEqualsBlank;
	}

	/** Sets whether explicit nulls may compare equal to blanks.
	 *
	 * @param nullEqualsBlank true to treat null and blank as equivalent */
	public void setNullEqualsBlank(boolean nullEqualsBlank) {
		this.nullEqualsBlank = nullEqualsBlank;
	}

	/** Returns whether whitespace-only strings are considered blank. */
	public boolean isWhitespaceOnlyIsBlank() {
		return whitespaceOnlyIsBlank;
	}

	/** Controls whether whitespace-only strings are treated as blank.
	 *
	 * @param whitespaceOnlyIsBlank true to consider whitespace-only strings blank */
	public void setWhitespaceOnlyIsBlank(boolean whitespaceOnlyIsBlank) {
		this.whitespaceOnlyIsBlank = whitespaceOnlyIsBlank;
	}

	/** Returns whether a blank expected value skips the cell comparison. */
	public boolean isIgnoreExpectedBlank() {
		return ignoreExpectedBlank;
	}

	/** Enables skipping this cell comparison when the expected value is blank or null.
	 *
	 * @param ignoreExpectedBlank true to skip blank expected values */
	public void setIgnoreExpectedBlank(boolean ignoreExpectedBlank) {
		this.ignoreExpectedBlank = ignoreExpectedBlank;
	}

	/** Returns the decimal scale used to validate numeric values. */
	public int getScale() {
		return scale;
	}

	/** Sets the decimal scale used to validate numeric values.
	 *
	 * @param scale number of fractional decimal places, from 0 through 12
	 *
	 * @throws IllegalArgumentException if the scale is outside 0 through 12 */
	public void setScale(int scale) {
		if (scale < 0 || scale > 12)
			throw new IllegalArgumentException("scale must be between 0 and 12");
		this.scale = scale;
	}

	/** Returns the rounding mode used for decimal scale validation. */
	public RoundingMode getRoundingMode() {
		return roundingMode;
	}

	/** Sets the decimal rounding mode, using {@link RoundingMode#UNNECESSARY} for null.
	 *
	 * @param roundingMode rounding mode for scale validation */
	public void setRoundingMode(RoundingMode roundingMode) {
		this.roundingMode = roundingMode == null ? RoundingMode.UNNECESSARY : roundingMode;
	}

	/** Returns the allowed absolute difference for decimal comparisons. */
	public String getTolerance() {
		return tolerance;
	}

	/** Sets the allowed absolute decimal difference.
	 *
	 * @param tolerance non-negative decimal tolerance
	 *
	 * @throws IllegalArgumentException if the value is blank, invalid, or negative */
	public void setTolerance(String tolerance) {
		if (tolerance == null || tolerance.isBlank())
			throw new IllegalArgumentException("tolerance must be a non-negative decimal");
		java.math.BigDecimal parsed = new java.math.BigDecimal(tolerance);
		if (parsed.signum() < 0)
			throw new IllegalArgumentException("tolerance must be non-negative");
		this.tolerance = parsed.toPlainString();
	}

	/** Returns the configured text values recognized as explicit nulls. */
	public List<String> getNullTokens() {
		return nullTokens;
	}

	/** Replaces the configured explicit-null tokens with a defensive copy.
	 *
	 * @param nullTokens token strings, or null to clear the list */
	public void setNullTokens(List<String> nullTokens) {
		this.nullTokens = nullTokens == null ? new ArrayList<>() : new ArrayList<>(nullTokens);
	}
}
