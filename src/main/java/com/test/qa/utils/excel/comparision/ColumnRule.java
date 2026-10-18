package com.test.qa.utils.excel.comparision;

import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;

/**
 * Configuration for comparing one report column. Missing fields use safe
 * defaults.
 */
public class ColumnRule {
	public enum DataType {
		STRING, DECIMAL
	}

	private DataType type = DataType.STRING;
	private boolean trimWhitespace = true;
	private boolean caseSensitive = true;
	private boolean nullEqualsBlank = false;
	private boolean whitespaceOnlyIsBlank = true;
	private int scale = 2;
	private RoundingMode roundingMode = RoundingMode.UNNECESSARY;
	private String tolerance = "0";
	private List<String> nullTokens = new ArrayList<>();

	public DataType getType() {
		return type;
	}

	public void setType(DataType type) {
		this.type = type == null ? DataType.STRING : type;
	}

	public boolean isTrimWhitespace() {
		return trimWhitespace;
	}

	public void setTrimWhitespace(boolean trimWhitespace) {
		this.trimWhitespace = trimWhitespace;
	}

	public boolean isCaseSensitive() {
		return caseSensitive;
	}

	public void setCaseSensitive(boolean caseSensitive) {
		this.caseSensitive = caseSensitive;
	}

	public boolean isNullEqualsBlank() {
		return nullEqualsBlank;
	}

	public void setNullEqualsBlank(boolean nullEqualsBlank) {
		this.nullEqualsBlank = nullEqualsBlank;
	}

	public boolean isWhitespaceOnlyIsBlank() {
		return whitespaceOnlyIsBlank;
	}

	public void setWhitespaceOnlyIsBlank(boolean whitespaceOnlyIsBlank) {
		this.whitespaceOnlyIsBlank = whitespaceOnlyIsBlank;
	}

	public int getScale() {
		return scale;
	}

	public void setScale(int scale) {
		if (scale < 0 || scale > 12)
			throw new IllegalArgumentException("scale must be between 0 and 12");
		this.scale = scale;
	}

	public RoundingMode getRoundingMode() {
		return roundingMode;
	}

	public void setRoundingMode(RoundingMode roundingMode) {
		this.roundingMode = roundingMode == null ? RoundingMode.UNNECESSARY : roundingMode;
	}

	public String getTolerance() {
		return tolerance;
	}

	public void setTolerance(String tolerance) {
		if (tolerance == null || tolerance.isBlank())
			throw new IllegalArgumentException("tolerance must be a non-negative decimal");
		java.math.BigDecimal parsed = new java.math.BigDecimal(tolerance);
		if (parsed.signum() < 0)
			throw new IllegalArgumentException("tolerance must be non-negative");
		this.tolerance = parsed.toPlainString();
	}

	public List<String> getNullTokens() {
		return nullTokens;
	}

	public void setNullTokens(List<String> nullTokens) {
		this.nullTokens = nullTokens == null ? new ArrayList<>() : new ArrayList<>(nullTokens);
	}
}
