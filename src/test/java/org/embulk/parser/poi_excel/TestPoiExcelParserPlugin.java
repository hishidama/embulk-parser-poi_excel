package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

import java.net.URL;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.List;
import java.util.TimeZone;

import org.embulk.config.ConfigSource;
import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.embulk.spi.time.Timestamp;
import org.junit.Test;

public class TestPoiExcelParserPlugin {

	@Test
	public void test1() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 1);
			parser.set("default_timezone", "Asia/Tokyo");
			parser.addColumn("boolean", "boolean");
			parser.addColumn("long", "long");
			parser.addColumn("double", "double");
			parser.addColumn("string", "string");
			parser.addColumn("timestamp", "timestamp").set("format", "%Y/%m/%d");

			URL inFile = getClass().getResource("test1.xls");
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

	@Test
	public void testRowNumber() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 1);
			parser.set("cell_error_null", false);
			parser.addColumn("sheet", "string").set("value", "sheet_name");
			parser.addColumn("sheet-n", "long").set("value", "sheet_name");
			parser.addColumn("row", "long").set("value", "row_number");
			parser.addColumn("flag", "boolean");
			parser.addColumn("col-n", "long").set("value", "column_number");
			parser.addColumn("col-s", "string").set("value", "column_number");

			URL inFile = getClass().getResource("test1.xls");
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

	@Test
	public void testForumlaReplace() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "formula_replace");

			ConfigSource replace0 = tester.newConfigSource();
			replace0.set("regex", "test1");
			replace0.set("to", "merged_cell");
			ConfigSource replace1 = tester.newConfigSource();
			replace1.set("regex", "B1");
			replace1.set("to", "B${row}");
			parser.set("formula_replace", Arrays.asList(replace0, replace1));

			parser.addColumn("text", "string");

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(2));
			assertThat(result.get(0).getAsString("text"), is("test3-a1"));
			assertThat(result.get(1).getAsString("text"), is("test2-b2"));
		}
	}

	@Test
	public void testSearchMergedCell_true() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "merged_cell");
			parser.addColumn("a", "string");
			parser.addColumn("b", "string");

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(4));
			check6(result, 0, "test3-a1", "test3-a1");
			check6(result, 1, "data", "0");
			check6(result, 2, null, null);
			check6(result, 3, null, null);
		}
	}

	@Test
	public void testSearchMergedCell_false() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "merged_cell");
			parser.set("search_merged_cell", false);
			parser.addColumn("a", "string");
			parser.addColumn("b", "string");

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(4));
			check6(result, 0, "test3-a1", null);
			check6(result, 1, "data", "0");
			check6(result, 2, null, null);
			check6(result, 3, null, null);
		}
	}

	private void check6(List<OutputRecord> result, int index, String a, String b) {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsString("a"), is(a));
		assertThat(r.getAsString("b"), is(b));
	}
}
