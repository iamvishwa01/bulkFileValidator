package com.test.qa.utils.excel.comparision;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.Closeable;
import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.EOFException;
import java.io.IOException;
import java.io.Writer;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.io.UncheckedIOException;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.PriorityQueue;
import java.util.UUID;
import java.util.function.Consumer;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

public class BulkFileComparator {
	private static final Object CONSOLE_LOCK = new Object();
	// ============================================================
	// CONFIG
	// ============================================================
	enum FileType {
		CSV, XLSX
	}
	public enum KeyType {
		STRING, NUMERIC
	}

	public static class Config {
		// Column used for ascending sort + row matching
		public String keyColumn = "ID";
		// NUMBER => 101 < 102 < 1000
		// STRING => lexical sorting
		public KeyType keyType = KeyType.STRING;
		// Number of records held in memory at a time
		public int chunkSize = 50_000;
		// Header comparison
		boolean caseInsensitiveHeaders = true;
		boolean trimHeaders = true;
		boolean strictHeaderOrder = false;
		// Value comparison
		boolean trimValues = true;
		boolean caseInsensitiveValues = false;
		// CSV
		char csvDelimiter = ',';
		// Report
		String reportSheetName = "Mismatches";
		String summarySheetName = "Summary";
	}

	// ============================================================
	// RECORD
	// ============================================================

	static class RecordData {
		String sortKey;
		String[] values;
		RecordData(String sortKey, String[] values) {
			this.sortKey = sortKey;
			this.values = values;
		}
	}

	// ============================================================
	// HEADER INFO
	// ============================================================

	static class HeaderInfo {
		List<String> headers;
		Map<String, Integer> indexMap;
		List<String> duplicateHeaders;
		HeaderInfo(List<String> headers, Config config) {
			this.headers = headers;
			this.indexMap = new LinkedHashMap<>();
			this.duplicateHeaders = new ArrayList<>();
			for (int i = 0; i < headers.size(); i++) {
				String normalized = normalizeHeader(headers.get(i), config);
				if (indexMap.containsKey(normalized)) {
					duplicateHeaders.add(headers.get(i));
				} else {
					indexMap.put(normalized, i);
				}
			}
		}

		Integer getIndex(String header, Config config) {

			return indexMap.get(normalizeHeader(header, config));
		}
	}

	// ============================================================
	// COLUMN MAPPING
	// ============================================================

	static class ColumnMapping {

		String columnName;
		int actualIndex;
		int expectedIndex;

		ColumnMapping(String columnName, int actualIndex, int expectedIndex) {

			this.columnName = columnName;
			this.actualIndex = actualIndex;
			this.expectedIndex = expectedIndex;
		}
	}

	// ============================================================
	// REPORT
	// ============================================================

	static class ReportWriter implements Closeable {

		private final SXSSFWorkbook workbook;
		private final Sheet summarySheet;
		private final Sheet mismatchSheet;
		private final Path htmlReport;
		private Writer htmlWriter;

		private int mismatchRow = 1;
		private boolean closed;
		// Defense-in-depth for reports: redact PAN-like digit sequences even when
		// source headers are poor or unknown. Header-based redaction is applied too.
		private static final java.util.regex.Pattern PAN_LIKE = java.util.regex.Pattern
				.compile("(?<!\\d)(?:\\d[ -]?){13,19}(?!\\d)");
		private static final java.util.regex.Pattern SENSITIVE_COLUMN = java.util.regex.Pattern.compile(
				"(?i).*(pan|card.?number|primary.?account|cvv|cvc|pin.?block|pin|track.?data|cryptogram|security.?code).*");

		long headerMismatches;
		long actualRows;
		long expectedRows;
		long valueMismatches;
		long missingRows;
		long extraRows;
		long duplicateKeys;
		String executionStatus = "COMPLETED";
		long malformedRows;

		ReportWriter(Config config, Path htmlReport) throws IOException {
			this.htmlReport = htmlReport;

			workbook = new SXSSFWorkbook(200);
			workbook.setCompressTempFiles(true);

			summarySheet = workbook.createSheet(config.summarySheetName);

			mismatchSheet = workbook.createSheet(config.reportSheetName);

			writeMismatchHeader();
			if (htmlReport != null) {
				Path parent = htmlReport.toAbsolutePath().getParent();
				if (parent != null) Files.createDirectories(parent);
				htmlWriter = Files.newBufferedWriter(htmlReport, StandardCharsets.UTF_8);
				writeHtml("<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>File validation report</title><style> :root{color-scheme:light;--ink:#172b4d;--muted:#5e6c84;--line:#e5eaf1;--surface:#fff;--bg:#f4f7fb;--blue:#2457c5;--red:#b42318;--green:#147d4a}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}main{max-width:1440px;margin:0 auto;padding:36px 32px 56px}h1{font-size:30px;line-height:1.2;margin:0 0 6px;letter-spacing:-.5px}.subtitle{color:var(--muted);margin:0 0 24px}.card{background:var(--surface);border:1px solid var(--line);border-radius:14px;box-shadow:0 3px 12px #172b4d0a;margin:20px 0;padding:24px}.result-card{display:flex;align-items:center;gap:16px;padding:20px 24px}.result-icon{width:44px;height:44px;border-radius:50%;display:grid;place-items:center;background:#fff0ef;color:var(--red);font-size:22px;font-weight:800}.result-icon.pass{background:#e8f7ef;color:var(--green)}.status{font-size:20px;font-weight:750;color:var(--red)}.status.pass{color:var(--green)}.result-note{color:var(--muted);font-size:13px;margin-top:2px}.section-heading{display:flex;align-items:baseline;justify-content:space-between;gap:12px;margin:0 0 18px}.section-heading h2{font-size:18px;margin:0}.section-heading p{color:var(--muted);font-size:13px;margin:0}.metrics{display:grid;grid-template-columns:repeat(5,minmax(130px,1fr));gap:12px}.metric{border:1px solid var(--line);border-radius:10px;padding:14px 16px;background:#fff}.metric-label{font-size:12px;color:var(--muted);font-weight:650;letter-spacing:.02em}.metric-value{font-size:24px;font-weight:750;margin-top:3px;font-variant-numeric:tabular-nums}.files{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:16px}.file{padding:12px 14px;background:#f8faff;border-radius:9px;min-width:0}.file-label{display:block;color:var(--muted);font-size:12px;font-weight:650;margin-bottom:3px}.file-value{overflow-wrap:anywhere;font-size:13px} .table-wrap{overflow:auto;border:1px solid var(--line);border-radius:10px}table{border-collapse:separate;border-spacing:0;width:100%;min-width:920px;background:white}th,td{text-align:left;padding:12px 14px;border-bottom:1px solid var(--line);vertical-align:top}th{position:sticky;top:0;background:#f5f7fb;color:#53627a;font-size:11px;text-transform:uppercase;letter-spacing:.055em;font-weight:750;white-space:nowrap}td{font-size:13px;line-height:1.45;overflow-wrap:anywhere}tbody tr:last-child td{border-bottom:0}tbody tr:nth-child(even){background:#fbfcfe}tbody tr:hover{background:#f4f8ff}td:nth-child(1){white-space:nowrap;font-weight:650;color:#344563}td:nth-child(2){min-width:130px}td:nth-child(3){min-width:130px}td:nth-child(4),td:nth-child(5){min-width:150px;max-width:280px}td:nth-child(6){min-width:280px;max-width:480px}.empty{padding:32px;text-align:center;color:var(--muted)}@media(max-width:900px){main{padding:24px 16px}.metrics{grid-template-columns:repeat(2,minmax(0,1fr))}.files{grid-template-columns:1fr}.card{padding:18px}}@media print{body{background:#fff}.card{box-shadow:none;break-inside:avoid}main{max-width:none;padding:0}.table-wrap{overflow:visible}table{min-width:0}th{position:static}}</style></head><body><main><h1>File validation report</h1><p class=\"subtitle\">Bulk file comparison results</p><section class=\"card result-card\"><div class=\"result-icon\" id=\"result-icon\">!</div><div><div id=\"result\" class=\"status\">IN PROGRESS</div><div class=\"result-note\">Comparison finished with the summary below.</div></div></section><section class=\"card\"><div class=\"section-heading\"><h2>Summary</h2><p>Validation totals and source files</p></div><div class=\"metrics\"><div class=\"metric\"><div class=\"metric-label\">Actual rows</div><div class=\"metric-value\" id=\"actual-rows\">–</div></div><div class=\"metric\"><div class=\"metric-label\">Expected rows</div><div class=\"metric-value\" id=\"expected-rows\">–</div></div><div class=\"metric\"><div class=\"metric-label\">Value mismatches</div><div class=\"metric-value\" id=\"value-mismatches\">–</div></div><div class=\"metric\"><div class=\"metric-label\">Missing / extra rows</div><div class=\"metric-value\" id=\"row-mismatches\">–</div></div><div class=\"metric\"><div class=\"metric-label\">Duplicate keys</div><div class=\"metric-value\" id=\"duplicate-keys\">–</div></div></div><div class=\"files\"><div class=\"file\"><span class=\"file-label\">Actual file</span><div id=\"actual-file\" class=\"file-value\"></div></div><div class=\"file\"><span class=\"file-label\">Expected file</span><div id=\"expected-file\" class=\"file-value\"></div></div></div><div class=\"files\"><div class=\"file\"><span class=\"file-label\">Header mismatches</span><div id=\"header-mismatches\" class=\"file-value\">–</div></div><div class=\"file\"><span class=\"file-label\">Malformed rows · execution status</span><div class=\"file-value\"><span id=\"malformed-rows\">–</span> · <span id=\"execution-status\">–</span></div></div></div></section><section class=\"card\"><div class=\"section-heading\"><h2>Mismatches</h2><p>Cell values are redacted where sensitive data is detected.</p></div><div class=\"table-wrap\"><table><thead><tr><th>Type</th><th>Key</th><th>Column</th><th>Actual</th><th>Expected</th><th>Details</th></tr></thead><tbody>");
			}
		}

		private void writeMismatchHeader() {

			Row row = mismatchSheet.createRow(0);

			row.createCell(0).setCellValue("Type");
			row.createCell(1).setCellValue("Key");
			row.createCell(2).setCellValue("Column");
			row.createCell(3).setCellValue("Actual");
			row.createCell(4).setCellValue("Expected");
			row.createCell(5).setCellValue("Details");
		}

		void addMismatch(String type, String key, String column, String actual, String expected, String details) {
			String safeType = safe(type), safeKey = redact(key, "key"), safeColumn = safe(column);
			String safeActual = redact(actual, column), safeExpected = redact(expected, column), safeDetails = redact(details, column);

			Row row = mismatchSheet.createRow(mismatchRow++);

			row.createCell(0).setCellValue(safeType);
			row.createCell(1).setCellValue(safeKey);
			row.createCell(2).setCellValue(safeColumn);
			row.createCell(3).setCellValue(safeActual);
			row.createCell(4).setCellValue(safeExpected);
			row.createCell(5).setCellValue(safeDetails);
			if (htmlWriter != null) {
				try {
				writeHtml("<tr><td>" + html(safeType) + "</td><td>" + html(safeKey) + "</td><td>" + html(safeColumn) + "</td><td>" + html(safeActual) + "</td><td>" + html(safeExpected) + "</td><td>" + html(safeDetails) + "</td></tr>");
				} catch (IOException e) { throw new UncheckedIOException(e); }
			}
		}

		void writeHeaderMismatch(String key, String column, String actual, String expected, String details) {

			headerMismatches++;

			addMismatch("HEADER_MISMATCH", key, column, actual, expected, details);
		}

		void writeSummary(Path actual, Path expected, boolean passed) {
			if (htmlWriter != null) {
				try {
					if (mismatchRow == 1) writeHtml("<tr><td colspan=\"6\" class=\"empty\">No mismatches found. Both files passed validation.</td></tr>");
					writeHtml("</tbody></table></div></section><script>const set=(id,value)=>document.getElementById(id).textContent=value;set('result','" + (passed ? "PASS" : "FAIL") + "');document.getElementById('result').className='status " + (passed ? "pass" : "fail") + "';set('result-icon','" + (passed ? "✓" : "!") + "');if(" + passed + ")document.getElementById('result-icon').classList.add('pass');set('actual-rows','" + actualRows + "');set('expected-rows','" + expectedRows + "');set('value-mismatches','" + valueMismatches + "');set('row-mismatches','" + (missingRows + extraRows) + "');set('duplicate-keys','" + duplicateKeys + "');set('header-mismatches','" + headerMismatches + "');set('malformed-rows','" + malformedRows + "');set('execution-status','" + js(executionStatus) + "');set('actual-file','" + js(actual.toAbsolutePath().toString()) + "');set('expected-file','" + js(expected.toAbsolutePath().toString()) + "');</script></main></body></html>");
					htmlWriter.close(); htmlWriter = null;
				} catch (IOException e) { throw new UncheckedIOException(e); }
			}

			summarySheet.createRow(0).createCell(0).setCellValue("File Validation Summary");

			writeSummaryRow(2, "Actual File", actual.toAbsolutePath().toString());

			writeSummaryRow(3, "Expected File", expected.toAbsolutePath().toString());

			writeSummaryRow(5, "Actual Rows", actualRows);
			writeSummaryRow(6, "Expected Rows", expectedRows);
			writeSummaryRow(7, "Value Mismatches", valueMismatches);
			writeSummaryRow(8, "Missing Rows", missingRows);
			writeSummaryRow(9, "Extra Rows", extraRows);
			writeSummaryRow(10, "Duplicate Keys", duplicateKeys);
			writeSummaryRow(11, "Header Mismatches", headerMismatches);
			writeSummaryRow(12, "Malformed Rows", malformedRows);
			writeSummaryRow(13, "Execution Status", executionStatus);

			writeSummaryRow(15, "Final Result", passed ? "PASS" : "FAIL");
		}

		private void writeSummaryRow(int rowNumber, String key, String value) {

			Row row = summarySheet.createRow(rowNumber);

			row.createCell(0).setCellValue(key);
			row.createCell(1).setCellValue(value);
		}

		private void writeSummaryRow(int rowNumber, String key, long value) {

			Row row = summarySheet.createRow(rowNumber);

			row.createCell(0).setCellValue(key);
			row.createCell(1).setCellValue(value);
		}

		void save(Path output) throws IOException {

			Files.createDirectories(output.toAbsolutePath().getParent());

			try (OutputStream os = Files.newOutputStream(output)) {
				workbook.write(os);
			} finally {
				close();
			}
		}

		private void writeHtml(String text) throws IOException { htmlWriter.write(text); }
		private static String js(String value) {
			return value.replace("\\", "\\\\").replace("'", "\\'").replace("\r", "\\r")
					.replace("\n", "\\n").replace("<", "\\u003c").replace(">", "\\u003e")
					.replace("&", "\\u0026");
		}
		private static String html(String value) {
			return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
					.replace("\"", "&quot;").replace("'", "&#39;");
		}

		private String safe(String value) {
			return value == null ? "" : value;
		}

		private String redact(String value, String columnName) {
			if (value == null)
				return "";
			String column = columnName == null ? "" : columnName;
			if (SENSITIVE_COLUMN.matcher(column).matches())
				return "[REDACTED]";
			return PAN_LIKE.matcher(value).replaceAll("[REDACTED_CARD_NUMBER]");
		}

		@Override
		public void close() {
			if (closed)
				return;
			closed = true;
			if (htmlWriter != null) {
				try { writeHtml("</tbody></table></div></section></main></body></html>"); htmlWriter.close(); }
				catch (IOException ignored) { }
				htmlWriter = null;
			}
			try {
				workbook.dispose();
			} finally {
				try {
					workbook.close();
				} catch (IOException ignored) {
				}
			}
		}

		public void printConsoleSummary(Path actual, Path expected, boolean passed) {

			synchronized (CONSOLE_LOCK) {

				System.out.println();
				System.out.println("==============================================");
				System.out.println("          FILE VALIDATION SUMMARY");
				System.out.println("==============================================");

				System.out.printf("%-25s : %s%n", "Actual File", actual.toAbsolutePath());

				System.out.printf("%-25s : %s%n", "Expected File", expected.toAbsolutePath());

				System.out.println("----------------------------------------------");

				System.out.printf("%-25s : %d%n", "Actual Rows", actualRows);

				System.out.printf("%-25s : %d%n", "Expected Rows", expectedRows);

				System.out.printf("%-25s : %d%n", "Value Mismatches", valueMismatches);

				System.out.printf("%-25s : %d%n", "Missing Rows", missingRows);

				System.out.printf("%-25s : %d%n", "Extra Rows", extraRows);

				System.out.printf("%-25s : %d%n", "Duplicate Keys", duplicateKeys);

				System.out.printf("%-25s : %d%n", "Header Mismatches", headerMismatches);
				System.out.printf("%-25s : %d%n", "Malformed Rows", malformedRows);

				System.out.println("----------------------------------------------");

				System.out.printf("%-25s : %s%n", "Final Result", passed ? "PASS" : "FAIL");

				System.out.println("==============================================");
				System.out.println();
			}
		}
	}

	// ============================================================
	// BINARY RUN WRITER
	// ============================================================

	static class RunWriter implements Closeable {

		private final DataOutputStream out;

		RunWriter(Path file) throws IOException {

			out = new DataOutputStream(new BufferedOutputStream(
					Files.newOutputStream(file, StandardOpenOption.CREATE_NEW, StandardOpenOption.WRITE)));
		}

		void write(RecordData record) throws IOException {

			writeString(record.sortKey);

			out.writeInt(record.values.length);

			for (String value : record.values) {
				writeString(value);
			}
		}

		private void writeString(String value) throws IOException {

			if (value == null) {
				out.writeInt(-1);
				return;
			}

			byte[] bytes = value.getBytes(StandardCharsets.UTF_8);

			out.writeInt(bytes.length);
			out.write(bytes);
		}

		@Override
		public void close() throws IOException {
			out.close();
		}
	}

	// ============================================================
	// BINARY RUN READER
	// ============================================================

	static class RunReader implements Closeable {

		private final DataInputStream in;

		RunReader(Path file) throws IOException {

			in = new DataInputStream(new BufferedInputStream(Files.newInputStream(file)));
		}

		RecordData next() throws IOException {

			int firstByte = in.read();
			if (firstByte == -1) {
				return null;
			}
			int secondByte = in.read();
			int thirdByte = in.read();
			int fourthByte = in.read();
			if ((secondByte | thirdByte | fourthByte) < 0) {
				throw new IOException("Truncated temporary sort run string length");
			}
			int keyLength = (firstByte << 24) | (secondByte << 16) | (thirdByte << 8) | fourthByte;
			String key = readString(keyLength);

			int columnCount;
			try {
				columnCount = in.readInt();
			} catch (EOFException e) {
				throw new IOException("Truncated temporary sort run column count", e);
			}
			if (columnCount < 0 || columnCount > 16_384) {
				throw new IOException("Invalid temporary sort run column count: " + columnCount);
			}

			String[] values = new String[columnCount];

			for (int i = 0; i < columnCount; i++) {
				values[i] = readString();
			}

			return new RecordData(key, values);
		}

		private String readString() throws IOException {

			int length;
			try {
				length = in.readInt();
			} catch (EOFException e) {
				throw new IOException("Truncated temporary sort run string length", e);
			}
			return readString(length);
		}

		private String readString(int length) throws IOException {

			if (length == -1) {
				return null;
			}

			if (length < 0 || length > 256 * 1024 * 1024) {
				throw new IOException("Invalid temporary sort run string length: " + length);
			}
			byte[] bytes = new byte[length];

			in.readFully(bytes);

			return new String(bytes, StandardCharsets.UTF_8);
		}

		@Override
		public void close() throws IOException {
			in.close();
		}
	}

	// ============================================================
	// SORTED STREAM / K-WAY MERGE
	// ============================================================

	static class SortedStream implements Closeable {

		static class RunHead {

			int index;
			RunReader reader;
			RecordData record;

			RunHead(int index, RunReader reader, RecordData record) {

				this.index = index;
				this.reader = reader;
				this.record = record;
			}
		}

		private final List<RunReader> readers;
		private final PriorityQueue<RunHead> queue;
		private final Config config;

		SortedStream(List<Path> runs, Config config) throws IOException {

			this.config = config;

			readers = new ArrayList<>();

			queue = new PriorityQueue<>((a, b) -> {

				int cmp = compareKeys(a.record.sortKey, b.record.sortKey, config);

				if (cmp != 0) {
					return cmp;
				}

				return Integer.compare(a.index, b.index);
			});

			for (int i = 0; i < runs.size(); i++) {

				RunReader reader = new RunReader(runs.get(i));

				readers.add(reader);

				RecordData record = reader.next();

				if (record != null) {

					queue.add(new RunHead(i, reader, record));
				}
			}
		}

		private RecordData buffered;

		RecordData peek() throws IOException {
			if (buffered == null) {
				buffered = takeNext();
			}
			return buffered;
		}

		RecordData next() throws IOException {
			if (buffered != null) {
				RecordData result = buffered;
				buffered = null;
				return result;
			}
			return takeNext();
		}

		private RecordData takeNext() throws IOException {
			RunHead head = queue.poll();
			if (head == null) {
				return null;
			}

			RecordData result = head.record;
			RecordData next = head.reader.next();
			if (next != null) {
				head.record = next;
				queue.add(head);
			}
			return result;
		}

		@Override
		public void close() throws IOException {

			for (RunReader reader : readers) {
				reader.close();
			}
		}
	}

	// ============================================================
	// EXCEL HEADER HANDLER
	// ============================================================

	static class StopParsingException extends SAXException {
	}

	static class ExcelHeaderHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

		private final Map<Integer, String> cells = new HashMap<>();

		List<String> header;

		@Override
		public void startRow(int rowNum) {
			cells.clear();
		}

		@Override
		public void endRow(int rowNum) {

			if (rowNum != 0) {
				return;
			}

			int max = cells.keySet().stream().max(Integer::compareTo).orElse(-1);

			List<String> result = new ArrayList<>();

			for (int i = 0; i <= max; i++) {

				result.add(cells.getOrDefault(i, ""));
			}

			header = result;

			throw new RuntimeException(new StopParsingException());
		}

		@Override
		public void cell(String cellReference, String formattedValue, XSSFComment comment) {

			int column = new CellReference(cellReference).getCol();

			cells.put(column, formattedValue);
		}

		@Override
		public void headerFooter(String text, boolean isHeader, String tagName) {
		}
	}

	// ============================================================
	// EXCEL DATA HANDLER
	// ============================================================

	static class ExcelDataHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

		private final int headerSize;

		private final Consumer<String[]> consumer;

		private final Map<Integer, String> cells = new HashMap<>();

		private boolean hasCell;

		ExcelDataHandler(int headerSize, Consumer<String[]> consumer) {

			this.headerSize = headerSize;
			this.consumer = consumer;
		}

		@Override
		public void startRow(int rowNum) {

			cells.clear();
			hasCell = false;
		}

		@Override
		public void endRow(int rowNum) {

			// Row 0 = header
			if (rowNum == 0) {
				return;
			}

			// Ignore completely empty spreadsheet rows
			if (!hasCell) {
				return;
			}

			String[] row = new String[headerSize];

			for (int i = 0; i < headerSize; i++) {

				row[i] = cells.getOrDefault(i, "");
			}

			try {

				consumer.accept(row);

			} catch (Exception e) {

				throw new RuntimeException(e);
			}
		}

		@Override
		public void cell(String cellReference, String formattedValue, XSSFComment comment) {

			hasCell = true;

			int column = new CellReference(cellReference).getCol();

			if (column < headerSize) {

				cells.put(column, formattedValue);
			}
		}

		@Override
		public void headerFooter(String text, boolean isHeader, String tagName) {
		}
	}

	// ============================================================
	// PUBLIC COMPARE METHOD
	// ============================================================

	/** Backward-compatible comparison using the global Config value rules. */
	public static void compare(Path actualFile, Path expectedFile, Path reportFile, Config config) throws Exception {
		compare(actualFile, expectedFile, reportFile, config, null);
	}

	/** Compare using per-column rules from a named validation profile. */
	public static void compare(Path actualFile, Path expectedFile, Path reportFile, Config config,
			ValidationProfile profile) throws Exception {

		Objects.requireNonNull(actualFile, "actualFile");
		Objects.requireNonNull(expectedFile, "expectedFile");
		Objects.requireNonNull(reportFile, "reportFile");
		Objects.requireNonNull(config, "config");
		Path actualPath = actualFile.toAbsolutePath().normalize();
		Path expectedPath = expectedFile.toAbsolutePath().normalize();
		Path reportPath = reportFile.toAbsolutePath().normalize();
		if (reportPath.equals(actualPath) || reportPath.equals(expectedPath)) {
			throw new IllegalArgumentException("Report path must be different from both input files");
		}
		if (!Files.isRegularFile(actualPath) || !Files.isReadable(actualPath)) {
			throw new IllegalArgumentException("Actual input must be a readable regular file: " + actualFile);
		}
		if (!Files.isRegularFile(expectedPath) || !Files.isReadable(expectedPath)) {
			throw new IllegalArgumentException("Expected input must be a readable regular file: " + expectedFile);
		}

		// Do not mutate the caller's Config: comparisons may run concurrently.
		Config effectiveConfig = copyConfig(config);
		if (profile != null) {
			if (profile.getProfileName() == null || profile.getProfileName().isBlank())
				throw new IllegalArgumentException("Validation profile profileName is required");
			if (profile.getKeyColumn() == null || profile.getKeyColumn().isBlank())
				throw new IllegalArgumentException("Validation profile keyColumn is required");
			effectiveConfig.keyColumn = profile.getKeyColumn();
		}
		config = effectiveConfig;

		String comparisonId = createComparisonId();

		System.out.println();
		System.out.println("[" + comparisonId + "] Starting comparison");

		System.out.println("[" + comparisonId + "] Actual   : " + actualFile.toAbsolutePath());

		System.out.println("[" + comparisonId + "] Expected : " + expectedFile.toAbsolutePath());

		Path htmlReport = htmlReportPath(reportPath);
		if (htmlReport != null && (htmlReport.equals(actualPath) || htmlReport.equals(expectedPath) || htmlReport.equals(reportPath))) {
			throw new IllegalArgumentException("HTML report path must be different from input files and XLSX report");
		}
		ReportWriter report = new ReportWriter(config, htmlReport);
		Path actualTemp = null;
		Path expectedTemp = null;
		boolean reportSaved = false;
		try {
			// Every invocation gets its own temporary directories, cleaned even when setup fails.
			actualTemp = Files.createTempDirectory("actual-sort-" + comparisonId + "-");
			expectedTemp = Files.createTempDirectory("expected-sort-" + comparisonId + "-");

			// ========================================================
			// READ HEADERS
			// ========================================================

			HeaderInfo actualHeader = readHeader(actualFile, config);

			HeaderInfo expectedHeader = readHeader(expectedFile, config);

			validateHeaders(actualHeader, expectedHeader, config, report);

			// ========================================================
			// DUPLICATE HEADER VALIDATION
			// ========================================================

			if (!actualHeader.duplicateHeaders.isEmpty() || !expectedHeader.duplicateHeaders.isEmpty()) {

				report.executionStatus = "NOT_EVALUATED_SCHEMA_FAILURE";
				report.writeSummary(actualFile, expectedFile, false);

				report.printConsoleSummary(actualFile, expectedFile, false);

				report.save(reportFile);
				reportSaved = true;

				return;
			}

			// ========================================================
			// KEY COLUMN
			// ========================================================

			Integer actualKeyIndex = actualHeader.getIndex(config.keyColumn, config);

			Integer expectedKeyIndex = expectedHeader.getIndex(config.keyColumn, config);

			if (actualKeyIndex == null) {

				report.writeHeaderMismatch("", config.keyColumn, "MISSING", "PRESENT",
						"Key column missing in actual file");
			}

			if (expectedKeyIndex == null) {

				report.writeHeaderMismatch("", config.keyColumn, "PRESENT", "MISSING",
						"Key column missing in expected file");
			}

			if (actualKeyIndex == null || expectedKeyIndex == null) {

				report.executionStatus = "NOT_EVALUATED_SCHEMA_FAILURE";
				report.writeSummary(actualFile, expectedFile, false);

				report.printConsoleSummary(actualFile, expectedFile, false);

				report.save(reportFile);
				reportSaved = true;

				return;
			}

			// ========================================================
			// COLUMN MAPPING
			// ========================================================

			List<ColumnMapping> mappings = createMappings(actualHeader, expectedHeader, config);

			// ========================================================
			// EXTERNAL SORT
			// ========================================================

			System.out.println("[" + comparisonId + "] Sorting actual file...");

			List<Path> actualRuns = createSortedRuns(actualFile, actualKeyIndex, actualHeader.headers.size(), "actual",
					actualTemp, config, report);

			System.out.println("[" + comparisonId + "] Actual sort completed. Runs = " + actualRuns.size());

			System.out.println("[" + comparisonId + "] Sorting expected file...");

			List<Path> expectedRuns = createSortedRuns(expectedFile, expectedKeyIndex, expectedHeader.headers.size(),
					"expected", expectedTemp, config, report);

			System.out.println("[" + comparisonId + "] Expected sort completed. Runs = " + expectedRuns.size());

			// ========================================================
			// STREAMING MERGE + COMPARISON
			// ========================================================

			System.out.println("[" + comparisonId + "] Comparing files...");

			try (SortedStream actualStream = new SortedStream(actualRuns, config);

					SortedStream expectedStream = new SortedStream(expectedRuns, config)) {

				compareSortedStreams(actualStream, expectedStream, mappings, config, profile, report);
			}

			// ========================================================
			// RESULT
			// ========================================================

			boolean passed = report.headerMismatches == 0 && report.valueMismatches == 0 && report.missingRows == 0
					&& report.extraRows == 0 && report.duplicateKeys == 0 && report.malformedRows == 0;

			report.writeSummary(actualFile, expectedFile, passed);

			report.printConsoleSummary(actualFile, expectedFile, passed);

			report.save(reportFile);
			reportSaved = true;

			System.out.println("[" + comparisonId + "] Report created: " + reportFile.toAbsolutePath());
			if (htmlReport != null) System.out.println("[" + comparisonId + "] HTML report created: " + htmlReport);

			System.out.println("[" + comparisonId + "] Comparison completed");

		} catch (Exception ex) {
			report.executionStatus = "ABORTED_PROCESSING_ERROR";
			if (!reportSaved) {
				try {
					report.addMismatch("PROCESSING_ERROR", "", "", "", "",
							"Comparison aborted before completion. See application logs for diagnostic details.");
					report.writeSummary(actualFile, expectedFile, false);
					report.save(reportFile);
					reportSaved = true;
				} catch (Exception reportFailure) {
					ex.addSuppressed(reportFailure);
				}
			}
			throw ex;
		} finally {

			// ========================================================
			// CLEANUP
			// ========================================================

			report.close();
			deleteDirectory(actualTemp);
			deleteDirectory(expectedTemp);

			System.out.println("[" + comparisonId + "] Temporary files cleaned");
		}
	}

	// ============================================================
	// HEADER VALIDATION
	// ============================================================

	private static void validateHeaders(HeaderInfo actual, HeaderInfo expected, Config config, ReportWriter report) {

		// Duplicate actual headers
		for (String header : actual.duplicateHeaders) {

			report.writeHeaderMismatch("", header, "DUPLICATE", "", "Duplicate header in actual");
		}

		// Duplicate expected headers
		for (String header : expected.duplicateHeaders) {

			report.writeHeaderMismatch("", header, "", "DUPLICATE", "Duplicate header in expected");
		}

		// Missing / extra headers
		for (String expectedHeader : expected.headers) {

			String normalized = normalizeHeader(expectedHeader, config);

			if (!actual.indexMap.containsKey(normalized)) {

				report.writeHeaderMismatch("", expectedHeader, "MISSING", "PRESENT", "Header missing in actual");
			}
		}

		for (String actualHeader : actual.headers) {

			String normalized = normalizeHeader(actualHeader, config);

			if (!expected.indexMap.containsKey(normalized)) {

				report.writeHeaderMismatch("", actualHeader, "EXTRA", "MISSING", "Extra header in actual");
			}
		}

		// Optional strict order validation
		if (config.strictHeaderOrder) {

			int min = Math.min(actual.headers.size(), expected.headers.size());

			for (int i = 0; i < min; i++) {

				String a = normalizeHeader(actual.headers.get(i), config);

				String e = normalizeHeader(expected.headers.get(i), config);

				if (!a.equals(e)) {

					report.writeHeaderMismatch("", "Column " + (i + 1), actual.headers.get(i), expected.headers.get(i),
							"Header order mismatch");
				}
			}
		}
	}

	// ============================================================
	// CREATE COLUMN MAPPING
	// ============================================================

	private static List<ColumnMapping> createMappings(HeaderInfo actual, HeaderInfo expected, Config config) {

		List<ColumnMapping> mappings = new ArrayList<>();

		for (int expectedIndex = 0; expectedIndex < expected.headers.size(); expectedIndex++) {

			String expectedHeader = expected.headers.get(expectedIndex);

			Integer actualIndex = actual.getIndex(expectedHeader, config);

			if (actualIndex != null) {

				mappings.add(new ColumnMapping(expectedHeader, actualIndex, expectedIndex));
			}
		}

		return mappings;
	}

	// ============================================================
	// CREATE SORTED RUNS
	// ============================================================

	private static List<Path> createSortedRuns(Path file, int keyIndex, int expectedColumnCount, String fileSide,
			Path tempDir, Config config, ReportWriter report) throws Exception {

		List<Path> runs = new ArrayList<>();

		List<RecordData> chunk = new ArrayList<>(config.chunkSize);

		AtomicCounter runCounter = new AtomicCounter();
		long[] recordNumber = { 1 };

		Consumer<String[]> addRow = row -> {
			long currentRecordNumber = ++recordNumber[0];

			if (row.length != expectedColumnCount) {
				report.malformedRows++;
				String rowKey = keyIndex < row.length ? normalizeValue(row[keyIndex], config) : "";
				report.addMismatch("MALFORMED_ROW", rowKey, "",
						"Actual field count: " + row.length, "Expected field count: " + expectedColumnCount,
						"Malformed " + fileSide + " record " + currentRecordNumber + " was skipped");
				return;
			}

			if (keyIndex >= row.length) {

				throw new RuntimeException("Key column index outside row: " + keyIndex);
			}

			String rawKey = row[keyIndex];

			String sortKey = normalizeValue(rawKey, config);

			if (sortKey == null || sortKey.isEmpty()) {
				report.malformedRows++;
				report.addMismatch("INVALID_KEY", "", config.keyColumn, rawKey, "",
						"Blank key in " + fileSide + " record " + currentRecordNumber + "; row was skipped");
				return;
			}

			chunk.add(new RecordData(sortKey, row));

			if (chunk.size() >= config.chunkSize) {

				try {

					Path run = writeSortedRun(chunk, tempDir, runCounter.next(), config);

					runs.add(run);

					chunk.clear();

				} catch (IOException e) {

					throw new UncheckedIOException(e);
				}
			}
		};

		FileType type = detectType(file);

		switch (type) {

		case CSV:
			streamCsv(file, config, addRow);
			break;

		case XLSX:
			streamExcel(file, config, addRow);
			break;

		default:
			throw new IllegalArgumentException("Unsupported file: " + file);
		}

		// Last chunk
		if (!chunk.isEmpty()) {

			Path run = writeSortedRun(chunk, tempDir, runCounter.next(), config);

			runs.add(run);
			chunk.clear();
		}

		return runs;
	}

	// ============================================================
	// WRITE SORTED RUN
	// ============================================================

	private static Path writeSortedRun(List<RecordData> records, Path tempDir, int runNumber, Config config)
			throws IOException {

		records.sort((a, b) -> compareKeys(a.sortKey, b.sortKey, config));

		Path run = tempDir.resolve(String.format("run-%06d.bin", runNumber));

		try (RunWriter writer = new RunWriter(run)) {

			for (RecordData record : records) {
				writer.write(record);
			}
		}

		return run;
	}

	// ============================================================
	// COMPARE SORTED STREAMS
	// ============================================================

	private static final class KeyGroup {
		final RecordData first;
		long count;

		KeyGroup(RecordData first) {
			this.first = first;
			this.count = 1;
		}
	}

	private static KeyGroup readKeyGroup(SortedStream stream, Config config) throws IOException {
		RecordData first = stream.next();
		if (first == null) {
			return null;
		}

		KeyGroup group = new KeyGroup(first);
		while (stream.peek() != null && compareKeys(first.sortKey, stream.peek().sortKey, config) == 0) {
			stream.next();
			group.count++;
		}
		return group;
	}

	private static void compareSortedStreams(SortedStream actualStream, SortedStream expectedStream,
			List<ColumnMapping> mappings, Config config, ValidationProfile profile, ReportWriter report)
			throws IOException {

		KeyGroup actual = readKeyGroup(actualStream, config);
		KeyGroup expected = readKeyGroup(expectedStream, config);

		while (actual != null || expected != null) {
			if (actual == null) {
				report.expectedRows += expected.count;
				report.missingRows += expected.count;
				if (expected.count > 1) {
					report.duplicateKeys++;
					report.addMismatch("DUPLICATE_KEY", expected.first.sortKey, "", "", "",
							"Duplicate key in expected file (" + expected.count + " occurrences)");
				}
				report.addMismatch("MISSING_ROW", expected.first.sortKey, "", "",
						String.join(" | ", expected.first.values),
						expected.count == 1 ? "Expected key not found in actual"
								: "Key group absent from actual; " + expected.count + " expected rows are unmatched");
				expected = readKeyGroup(expectedStream, config);
				continue;
			}

			if (expected == null) {
				report.actualRows += actual.count;
				report.extraRows += actual.count;
				if (actual.count > 1) {
					report.duplicateKeys++;
					report.addMismatch("DUPLICATE_KEY", actual.first.sortKey, "", "", "",
							"Duplicate key in actual file (" + actual.count + " occurrences)");
				}
				report.addMismatch("EXTRA_ROW", actual.first.sortKey, "", String.join(" | ", actual.first.values), "",
						actual.count == 1 ? "Actual key not found in expected"
								: "Key group absent from expected; " + actual.count + " actual rows are unmatched");
				actual = readKeyGroup(actualStream, config);
				continue;
			}

			int keyComparison = compareKeys(actual.first.sortKey, expected.first.sortKey, config);
			if (keyComparison < 0) {
				report.actualRows += actual.count;
				report.extraRows += actual.count;
				if (actual.count > 1) {
					report.duplicateKeys++;
					report.addMismatch("DUPLICATE_KEY", actual.first.sortKey, "", "", "",
							"Duplicate key in actual file (" + actual.count + " occurrences)");
				}
				report.addMismatch("EXTRA_ROW", actual.first.sortKey, "", String.join(" | ", actual.first.values), "",
						actual.count == 1 ? "Actual key not found in expected"
								: "Key group absent from expected; " + actual.count + " actual rows are unmatched");
				actual = readKeyGroup(actualStream, config);
				continue;
			}

			if (keyComparison > 0) {
				report.expectedRows += expected.count;
				report.missingRows += expected.count;
				if (expected.count > 1) {
					report.duplicateKeys++;
					report.addMismatch("DUPLICATE_KEY", expected.first.sortKey, "", "", "",
							"Duplicate key in expected file (" + expected.count + " occurrences)");
				}
				report.addMismatch("MISSING_ROW", expected.first.sortKey, "", "",
						String.join(" | ", expected.first.values),
						expected.count == 1 ? "Expected key not found in actual"
								: "Key group absent from actual; " + expected.count + " expected rows are unmatched");
				expected = readKeyGroup(expectedStream, config);
				continue;
			}

			// Same key exists in both files. Duplicate groups are ambiguous: do not
			// pair rows by arbitrary sort/run order. Report each duplicate group once.
			report.actualRows += actual.count;
			report.expectedRows += expected.count;
			if (actual.count > 1 || expected.count > 1) {
				if (actual.count > 1) {
					report.duplicateKeys++;
					report.addMismatch("DUPLICATE_KEY", actual.first.sortKey, "", "", "",
							"Duplicate key in actual file (" + actual.count + " occurrences)");
				}
				if (expected.count > 1) {
					report.duplicateKeys++;
					report.addMismatch("DUPLICATE_KEY", expected.first.sortKey, "", "", "",
							"Duplicate key in expected file (" + expected.count + " occurrences)");
				}
				report.addMismatch("AMBIGUOUS_KEY_MATCH", actual.first.sortKey, "",
						"Actual occurrences: " + actual.count, "Expected occurrences: " + expected.count,
						"Rows were not paired because this key is duplicated in at least one file");
			} else {
				for (ColumnMapping mapping : mappings) {
					String actualValue = getValue(actual.first.values, mapping.actualIndex);
					String expectedValue = getValue(expected.first.values, mapping.expectedIndex);
					if (profile == null) {
						if (!valuesEqual(actualValue, expectedValue, config)) {
							report.valueMismatches++;
							report.addMismatch("VALUE_MISMATCH", actual.first.sortKey, mapping.columnName, actualValue,
									expectedValue, "Value mismatch under global Config settings");
						}
					} else {
						ColumnRule rule = profile.ruleFor(mapping.columnName);
						RuleBasedValueComparator.ComparisonResult result = RuleBasedValueComparator.compare(actualValue,
								expectedValue, rule);
						if (!result.matches()) {
							report.valueMismatches++;
							report.addMismatch("VALUE_MISMATCH", actual.first.sortKey, mapping.columnName,
									result.normalizedActual(), result.normalizedExpected(), result.explanation());
						}
					}
				}
			}

			actual = readKeyGroup(actualStream, config);
			expected = readKeyGroup(expectedStream, config);
		}
	}

	// ============================================================
	// CSV STREAMING
	// ============================================================

	private static void streamCsv(Path file, Config config, Consumer<String[]> consumer) throws Exception {

		CSVFormat format = CSVFormat.DEFAULT.builder().setDelimiter(config.csvDelimiter).setIgnoreEmptyLines(false)
				.build();

		try (Reader reader = Files.newBufferedReader(file, StandardCharsets.UTF_8);

				CSVParser parser = format.parse(reader)) {

			Iterator<CSVRecord> iterator = parser.iterator();

			// Skip header
			if (iterator.hasNext()) {
				iterator.next();
			}

			while (iterator.hasNext()) {

				CSVRecord record = iterator.next();

				String[] row = new String[record.size()];

				for (int i = 0; i < record.size(); i++) {

					row[i] = record.get(i);
				}

				consumer.accept(row);
			}
		}
	}

	// ============================================================
	// EXCEL STREAMING
	// ============================================================

	private static void streamExcel(Path file, Config config, Consumer<String[]> consumer) throws Exception {

		HeaderInfo header = readExcelHeader(file, config);

		int headerSize = header.headers.size();

		try (OPCPackage pkg = OPCPackage.open(file.toFile())) {

			XSSFReader reader = new XSSFReader(pkg);

			StylesTable styles = reader.getStylesTable();

			SharedStrings strings = reader.getSharedStringsTable();

			XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();

			if (!sheets.hasNext()) {
				throw new IllegalArgumentException("No worksheet found: " + file);
			}

			try (InputStream sheet = sheets.next()) {

				XMLReader parser = XMLHelper.newXMLReader();

				ExcelDataHandler handler = new ExcelDataHandler(headerSize, consumer);

				parser.setContentHandler(
						new XSSFSheetXMLHandler(styles, null, strings, handler, new DataFormatter(), false));

				parser.parse(new InputSource(sheet));
			}
		}
	}

	// ============================================================
	// READ HEADER
	// ============================================================

	private static HeaderInfo readHeader(Path file, Config config) throws Exception {

		FileType type = detectType(file);

		List<String> headers;

		switch (type) {

		case CSV:

			headers = readCsvHeader(file, config);

			break;

		case XLSX:

			headers = readExcelHeader(file, config).headers;

			break;

		default:

			throw new IllegalArgumentException("Unsupported file type: " + file);
		}

		return new HeaderInfo(headers, config);
	}

	// ============================================================
	// CSV HEADER
	// ============================================================

	private static List<String> readCsvHeader(Path file, Config config) throws Exception {

		CSVFormat format = CSVFormat.DEFAULT.builder().setDelimiter(config.csvDelimiter).setIgnoreEmptyLines(false)
				.build();

		try (Reader reader = Files.newBufferedReader(file, StandardCharsets.UTF_8);

				CSVParser parser = format.parse(reader)) {

			Iterator<CSVRecord> iterator = parser.iterator();

			if (!iterator.hasNext()) {

				throw new IllegalArgumentException("CSV file is empty: " + file);
			}

			CSVRecord record = iterator.next();

			List<String> headers = new ArrayList<>();

			for (String value : record) {

				headers.add(value);
			}

			return headers;
		}
	}

	// ============================================================
	// EXCEL HEADER
	// ============================================================

	private static HeaderInfo readExcelHeader(Path file, Config config) throws Exception {

		try (OPCPackage pkg = OPCPackage.open(file.toFile())) {

			XSSFReader reader = new XSSFReader(pkg);

			StylesTable styles = reader.getStylesTable();

			SharedStrings strings = reader.getSharedStringsTable();

			XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();

			if (!sheets.hasNext()) {

				throw new IllegalArgumentException("No worksheet found: " + file);
			}

			try (InputStream sheet = sheets.next()) {

				XMLReader parser = XMLHelper.newXMLReader();

				ExcelHeaderHandler handler = new ExcelHeaderHandler();

				parser.setContentHandler(
						new XSSFSheetXMLHandler(styles, null, strings, handler, new DataFormatter(), false));

				try {

					parser.parse(new InputSource(sheet));

				} catch (RuntimeException e) {

					if (!(e.getCause() instanceof StopParsingException)) {

						throw e;
					}
				}

				if (handler.header == null) {

					throw new IllegalArgumentException("Could not read Excel header: " + file);
				}

				return new HeaderInfo(handler.header, config);
			}
		}
	}

	// ============================================================
	private static Config copyConfig(Config source) {
		Config copy = new Config();
		copy.keyColumn = source.keyColumn;
		copy.keyType = source.keyType;
		copy.chunkSize = source.chunkSize;
		copy.caseInsensitiveHeaders = source.caseInsensitiveHeaders;
		copy.trimHeaders = source.trimHeaders;
		copy.strictHeaderOrder = source.strictHeaderOrder;
		copy.trimValues = source.trimValues;
		copy.caseInsensitiveValues = source.caseInsensitiveValues;
		copy.csvDelimiter = source.csvDelimiter;
		copy.reportSheetName = source.reportSheetName;
		copy.summarySheetName = source.summarySheetName;
		if (copy.chunkSize < 1)
			throw new IllegalArgumentException("chunkSize must be greater than zero");
		if (copy.keyType == null)
			throw new IllegalArgumentException("keyType cannot be null");
		return copy;
	}

	// VALUE HELPERS
	// ============================================================

	private static String getValue(String[] values, int index) {

		if (index >= values.length) {
			return "";
		}

		return values[index] == null ? "" : values[index];
	}

	private static boolean valuesEqual(String actual, String expected, Config config) {

		String a = normalizeValue(actual, config);

		String e = normalizeValue(expected, config);

		return Objects.equals(a, e);
	}

	private static String normalizeValue(String value, Config config) {

		if (value == null) {
			return "";
		}

		String result = value;

		if (config.trimValues) {
			result = result.trim();
		}

		if (config.caseInsensitiveValues) {
			result = result.toLowerCase(Locale.ROOT);
		}

		return result;
	}

	private static String normalizeHeader(String header, Config config) {

		if (header == null) {
			return "";
		}

		String result = header.replace("\uFEFF", "");

		if (config.trimHeaders) {
			result = result.trim();
		}

		if (config.caseInsensitiveHeaders) {
			result = result.toLowerCase(Locale.ROOT);
		}

		return result;
	}

	// ============================================================
	// KEY COMPARISON
	// ============================================================

	private static int compareKeys(String a, String b, Config config) {

		if (config.keyType == KeyType.NUMERIC) {
			try {
				BigDecimal bdA = new BigDecimal(a);
				BigDecimal bdB = new BigDecimal(b);
				return bdA.compareTo(bdB);
			} catch (NumberFormatException ex) {
				throw new IllegalArgumentException(
						"Invalid numeric key; configure STRING for identifier keys or correct the source data. Values: '"
								+ a + "' and '" + b + "'",
						ex);
			}
		}

		return a.compareTo(b);
	}

	// ============================================================
	// FILE TYPE
	// ============================================================

	private static FileType detectType(Path file) {

		String name = file.getFileName().toString().toLowerCase(Locale.ROOT);

		if (name.endsWith(".csv")) {
			return FileType.CSV;
		}

		if (name.endsWith(".xlsx")) {
			return FileType.XLSX;
		}

		throw new IllegalArgumentException("Only CSV and XLSX are supported: " + file);
	}

	// ============================================================
	// DELETE TEMP DIRECTORY
	// ============================================================

	private static void deleteDirectory(Path directory) {

		if (directory == null || !Files.exists(directory)) {
			return;
		}

		try {

			try (java.util.stream.Stream<Path> paths = Files.walk(directory)) {
				paths.sorted(Comparator.reverseOrder()).forEach(path -> {
					try {
						Files.deleteIfExists(path);
					} catch (IOException ignored) {
						// Cleanup is best-effort; operational logging can be added by caller.
					}
				});
			}

		} catch (IOException ignored) {
		}
	}

	// ============================================================
	// SIMPLE COUNTER
	// ============================================================

	static class AtomicCounter {

		private int value = 0;

		synchronized int next() {
			return value++;
		}
	}

	private static String createComparisonId() {
		return UUID.randomUUID().toString().substring(0, 8);
	}

	private static Path htmlReportPath(Path reportFile) {
		String name = reportFile.getFileName().toString();
		int extension = name.lastIndexOf('.');
		String base = extension > 0 ? name.substring(0, extension) : name;
		return reportFile.resolveSibling(base + ".html").toAbsolutePath().normalize();
	}
}
