package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.MatcherAssert.assertThat;

import java.net.URL;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.List;
import java.util.TimeZone;

import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.embulk.spi.time.Timestamp;
import org.junit.experimental.theories.DataPoints;
import org.junit.experimental.theories.Theories;
import org.junit.experimental.theories.Theory;
import org.junit.runner.RunWith;

@RunWith(Theories.class)
public class TestPoiExcelParserPlugin {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void test1(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheets", Arrays.asList("test1"));
			parser.set("skip_header_lines", 1);
			parser.set("default_timezone", "Asia/Tokyo");
			parser.addColumn("boolean", "boolean");
			parser.addColumn("long", "long");
			parser.addColumn("double", "double");
			parser.addColumn("string", "string");
			parser.addColumn("timestamp", "timestamp").set("format", "%Y/%m/%d");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			check1(result, 0, true, 123L, 123.4d, "abc", "2015/10/4");
			check1(result, 1, false, 456L, 456.7d, "def", "2015/10/5");
			check1(result, 2, false, 123L, 123d, "456", "2015/10/6");
			check1(result, 3, true, 123L, 123.4d, "abc", "2015/10/7");
			check1(result, 4, true, 123L, 123.4d, "abc", "2015/10/4");
			check1(result, 5, true, 1L, 1d, "true", null);
			check1(result, 6, null, null, null, null, null);
		}
	}

	private SimpleDateFormat sdf;
	{
		sdf = new SimpleDateFormat("yyyy/MM/dd");
		sdf.setTimeZone(TimeZone.getTimeZone("Asia/Tokyo"));
	}

	private void check1(List<OutputRecord> result, int index, Boolean b, Long l, Double d, String s, String t)
			throws ParseException {
		Timestamp timestamp = (t != null) ? Timestamp.ofEpochMilli(sdf.parse(t).getTime()) : null;

		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsBoolean("boolean"), is(b));
		assertThat(r.getAsLong("long"), is(l));
		assertThat(r.getAsDouble("double"), is(d));
		assertThat(r.getAsString("string"), is(s));
		assertThat(r.getAsTimestamp("timestamp"), is(timestamp));
	}

	@Theory
	public void testNumricFormat(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheets", Arrays.asList("test1"));
			parser.set("skip_header_lines", 1);
			parser.addColumn("value", "string").set("column_number", "C").set("numeric_format", "%.2f");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			checkNumricFormat(result, 0, "123.40");
			checkNumricFormat(result, 1, "456.70");
			checkNumricFormat(result, 2, "123.00");
		}
	}

	private void checkNumricFormat(List<OutputRecord> result, int index, String s) {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsString("value"), is(s));
	}

	@Theory
	public void testRowNumber(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheets", Arrays.asList("test1"));
			parser.set("skip_header_lines", 1);
			parser.addColumn("sheet", "string").set("value", "sheet_name");
			parser.addColumn("sheet-n", "long").set("value", "sheet_name");
			parser.addColumn("row", "long").set("value", "row_number");
			parser.addColumn("flag", "boolean");
			parser.addColumn("col-n", "long").set("value", "column_number");
			parser.addColumn("col-s", "string").set("value", "column_number");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			check4(result, 0, "test1", true);
			check4(result, 1, "test1", false);
			check4(result, 2, "test1", false);
			check4(result, 3, "test1", true);
			check4(result, 4, "test1", true);
			check4(result, 5, "test1", true);
			check4(result, 6, "test1", null);
		}
	}

	private void check4(List<OutputRecord> result, int index, String sheetName, Boolean b) {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsString("sheet"), is(sheetName));
		assertThat(r.getAsLong("sheet-n"), is(0L));
		assertThat(r.getAsLong("row"), is((long) (index + 2)));
		assertThat(r.getAsBoolean("flag"), is(b));
		assertThat(r.getAsLong("col-n"), is(1L));
		assertThat(r.getAsString("col-s"), is("A"));
	}

	@Theory
	public void test_sheets(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheets", Arrays.asList("formula_replace", "merged_cell"));
			parser.addColumn("a", "string");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(2 + 4));
			assertThat(result.get(0).getAsString("a"), is("boolean"));
			assertThat(result.get(1).getAsString("a"), is("test2-b1"));
			assertThat(result.get(2).getAsString("a"), is("test3-a1"));
			assertThat(result.get(3).getAsString("a"), is("data"));
		}
	}
}
