package com.test.qa.utils.excel.comparision;

import java.util.LinkedHashMap;
import java.util.Map;

/** Named, reusable set of column rules for one customer/report type. */
public class ValidationProfile {
	private String profileName;
	private String keyColumn = "ID";
	private Map<String, ColumnRule> columns = new LinkedHashMap<>();

	/** Returns the profile name used to identify this validation configuration. */
	public String getProfileName() {
		return profileName;
	}

	/** Sets the profile name.
	 * @param profileName descriptive name for the profile */
	public void setProfileName(String profileName) {
		this.profileName = profileName;
	}

	/** Returns the column header used to pair records between input sources. */
	public String getKeyColumn() {
		return keyColumn;
	}

	/** Sets the record key column.
	 * @param keyColumn header name of the key column */
	public void setKeyColumn(String keyColumn) {
		this.keyColumn = keyColumn;
	}

	/** Returns all configured comparison rules, keyed by column header. */
	public Map<String, ColumnRule> getColumns() {
		return columns;
	}

	/** Replaces the configured rules with a defensive copy.
	 * @param columns rules keyed by header name */
	public void setColumns(Map<String, ColumnRule> columns) {
		this.columns = columns == null ? new LinkedHashMap<>() : new LinkedHashMap<>(columns);
	}

	/** Finds a rule by column header, ignoring header capitalization.
	 *
	 * @param columnName header to look up
	 *
	 * @return the configured rule, or a default rule when no rule matches */
	public ColumnRule ruleFor(String columnName) {
		ColumnRule rule = columns.get(columnName);
		if (rule != null)
			return rule;
		for (Map.Entry<String, ColumnRule> entry : columns.entrySet()) {
			if (entry.getKey().equalsIgnoreCase(columnName))
				return entry.getValue();
		}
		return new ColumnRule();
	}
}
