package com.test.qa.utils.excel.comparision;

import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.json.JsonMapper;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.deser.std.StdDeserializer;
import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.DeserializationContext;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.math.RoundingMode;

/** Parses and validates JSON-based column comparison profiles. */
public final class RuleConfigLoader {
	private static final ObjectMapper MAPPER = createMapper();

	/** Prevents construction of this static utility class. */
	private RuleConfigLoader() {
	}

	/** Creates the JSON mapper, including case-insensitive parsing of rounding modes. */
	private static ObjectMapper createMapper() {
		SimpleModule module = new SimpleModule();
		module.addDeserializer(RoundingMode.class, new StdDeserializer<>(RoundingMode.class) {
			/** Converts a JSON rounding mode name to its Java enum value. */
			@Override
			public RoundingMode deserialize(JsonParser parser, DeserializationContext context) throws IOException {
				try {
					return RoundingMode.valueOf(parser.getValueAsString().trim().toUpperCase(java.util.Locale.ROOT));
				} catch (RuntimeException ex) {
					throw new IOException("Unsupported roundingMode: " + parser.getValueAsString(), ex);
				}
			}
		});
		return JsonMapper.builder().addModule(module).configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, true)
				.build();
	}

	/** Loads and validates a profile from a filesystem path.
	 *
	 * @param path JSON profile file
	 *
	 * @return parsed validation profile
	 *
	 * @throws IOException if the file cannot be read or contains invalid configuration */
	public static ValidationProfile fromFile(Path path) throws IOException {
		try (InputStream in = Files.newInputStream(path)) {
			return validate(MAPPER.readValue(in, ValidationProfile.class));
		}
	}

	/** Loads and validates a profile from the application's classpath resources.
	 *
	 * @param resourcePath classpath resource name
	 *
	 * @return parsed validation profile
	 *
	 * @throws IOException if the resource is missing, unreadable, or invalid */
	public static ValidationProfile fromResource(String resourcePath) throws IOException {
		try (InputStream in = RuleConfigLoader.class.getClassLoader().getResourceAsStream(resourcePath)) {
			if (in == null)
				throw new IOException("Profile resource not found: " + resourcePath);
			return validate(MAPPER.readValue(in, ValidationProfile.class));
		}
	}

	/** Checks required profile fields and validates each rule's decimal tolerance. */
	private static ValidationProfile validate(ValidationProfile profile) throws IOException {
		if (profile == null)
			throw new IOException("Profile is empty");
		if (profile.getProfileName() == null || profile.getProfileName().isBlank())
			throw new IOException("profileName is required");
		if (profile.getKeyColumn() == null || profile.getKeyColumn().isBlank())
			throw new IOException("keyColumn is required");
		if (profile.getColumns() == null)
			throw new IOException("columns cannot be null");
		for (var entry : profile.getColumns().entrySet()) {
			if (entry.getKey() == null || entry.getKey().isBlank() || entry.getValue() == null)
				throw new IOException("Each column rule must have a column name and rule");
			ColumnRule rule = entry.getValue();
			if (rule.getType() == ColumnRule.DataType.DECIMAL && rule.getRoundingMode() == null)
				throw new IOException("roundingMode cannot be null for column " + entry.getKey());
			try {
				rule.setTolerance(rule.getTolerance());
			} catch (RuntimeException ex) {
				throw new IOException("Invalid tolerance for column " + entry.getKey(), ex);
			}
		}
		return profile;
	}
}
