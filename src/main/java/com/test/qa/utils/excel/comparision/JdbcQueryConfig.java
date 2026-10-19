package com.test.qa.utils.excel.comparision;

import java.util.Objects;

/** Holds credentials, SQL, and cursor settings for a database-backed expected source. */
public final class JdbcQueryConfig {
	/** Environment variable containing the JDBC connection URL. */
	public static final String JDBC_URL_ENV = "BULK_VALIDATOR_DB_URL";
	/** Environment variable containing the database username. */
	public static final String USERNAME_ENV = "BULK_VALIDATOR_DB_USER";
	/** Environment variable containing the database password. */
	public static final String PASSWORD_ENV = "BULK_VALIDATOR_DB_PASSWORD";

	private final String jdbcUrl;
	private final String username;
	private final String password;
	private final String query;
	private int fetchSize = 1_000;
	private int queryTimeoutSeconds;

	/** Creates a query source and validates its connection settings and SQL.
	 *
	 * @param jdbcUrl MySQL JDBC URL
	 *
	 * @param username database account name
	 *
	 * @param password database account password
	 *
	 * @param query read-only SQL query that supplies expected rows
	 *
	 * @throws IllegalArgumentException if required text is missing or the URL is not MySQL */
	public JdbcQueryConfig(String jdbcUrl, String username, String password, String query) {
		this.jdbcUrl = requireText(jdbcUrl, "jdbcUrl");
		if (!this.jdbcUrl.startsWith("jdbc:mysql:"))
			throw new IllegalArgumentException("Only MySQL JDBC URLs are currently supported");
		this.username = requireText(username, "username");
		this.password = Objects.requireNonNull(password, "password");
		String normalizedQuery = requireText(query, "query").trim();
		if (normalizedQuery.endsWith(";"))
			normalizedQuery = normalizedQuery.substring(0, normalizedQuery.length() - 1).trim();
		this.query = normalizedQuery;
	}

	/** Builds a source using credentials from the process environment and the supplied SQL.
	 *
	 * @param query read-only SQL query that supplies expected rows
	 *
	 * @return configured MySQL query source
	 *
	 * @throws IllegalStateException if a required connection variable is missing */
	public static JdbcQueryConfig fromEnvironment(String query) {
		return new JdbcQueryConfig(requiredEnvironment(JDBC_URL_ENV), requiredEnvironment(USERNAME_ENV),
				requiredEnvironment(PASSWORD_ENV), query);
	}

	/** Reads a required environment variable and rejects missing or blank values. */
	private static String requiredEnvironment(String name) {
		String value = System.getenv(name);
		if (value == null || value.isBlank())
			throw new IllegalStateException("Required environment variable is not set: " + name);
		return value;
	}

	/** Validates a required nonblank string and returns it unchanged. */
	private static String requireText(String value, String name) {
		if (value == null || value.isBlank())
			throw new IllegalArgumentException(name + " cannot be blank");
		return value;
	}

	/** Returns the configured JDBC URL. */
	public String getJdbcUrl() { return jdbcUrl; }
	/** Returns the configured database username. */
	public String getUsername() { return username; }
	/** Returns the configured database password for JDBC connection setup. */
	public String getPassword() { return password; }
	/** Returns the normalized SQL query. */
	public String getQuery() { return query; }
	/** Returns the requested JDBC fetch size. */
	public int getFetchSize() { return fetchSize; }
	/** Returns the query timeout in seconds; zero selects the driver default. */
	public int getQueryTimeoutSeconds() { return queryTimeoutSeconds; }

	/** Sets how many result rows the JDBC driver should fetch at a time.
	 *
	 * @param fetchSize positive number of rows
	 *
	 * @return this configuration for chained setup
	 *
	 * @throws IllegalArgumentException if fetchSize is less than one */
	public JdbcQueryConfig setFetchSize(int fetchSize) {
		if (fetchSize < 1) throw new IllegalArgumentException("fetchSize must be greater than zero");
		this.fetchSize = fetchSize;
		return this;
	}

	/** Sets the maximum time allowed for the query; zero uses the driver's default.
	 *
	 * @param queryTimeoutSeconds timeout in seconds, or zero for the driver default
	 *
	 * @return this configuration for chained setup
	 *
	 * @throws IllegalArgumentException if the timeout is negative */
	public JdbcQueryConfig setQueryTimeoutSeconds(int queryTimeoutSeconds) {
		if (queryTimeoutSeconds < 0) throw new IllegalArgumentException("queryTimeoutSeconds cannot be negative");
		this.queryTimeoutSeconds = queryTimeoutSeconds;
		return this;
	}

	/** Returns a safe source label without exposing connection credentials or SQL. */
	public String getReportDescription() {
		return "MySQL query result";
	}
}
