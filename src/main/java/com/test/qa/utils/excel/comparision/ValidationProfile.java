package com.test.qa.utils.excel.comparision;

import java.util.LinkedHashMap;
import java.util.Map;

/** Named, reusable set of column rules for one customer/report type. */
public class ValidationProfile {
	private String profileName;
	private String keyColumn = "ID";
	private Map<String, ColumnRule> columns = new LinkedHashMap<>();

	public String getProfileName() {
		return profileName;
	}

	public void setProfileName(String profileName) {
		this.profileName = profileName;
	}

	public String getKeyColumn() {
		return keyColumn;
	}

	public void setKeyColumn(String keyColumn) {
		this.keyColumn = keyColumn;
	}

	public Map<String, ColumnRule> getColumns() {
		return columns;
	}

	public void setColumns(Map<String, ColumnRule> columns) {
		this.columns = columns == null ? new LinkedHashMap<>() : new LinkedHashMap<>(columns);
	}

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
