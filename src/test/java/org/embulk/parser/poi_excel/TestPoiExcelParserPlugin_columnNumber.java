package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.hamcrest.MatcherAssert.assertThat;

import java.net.URL;
import java.text.ParseException;
import java.util.List;

import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.junit.experimental.theories.DataPoints;
import org.junit.experimental.theories.Theories;
import org.junit.experimental.theories.Theory;
import org.junit.runner.RunWith;

@RunWith(Theories.class)
public class TestPoiExcelParserPlugin_columnNumber {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void testColumnNumber_string(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 1);
			parser.addColumn("text", "string").set("column_number", "D");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			assertThat(result.get(0).getAsString("text"), is("abc"));
			assertThat(result.get(1).getAsString("text"), is("def"));
			assertThat(result.get(2).getAsString("text"), is("456"));
			assertThat(result.get(3).getAsString("text"), is("abc"));
			assertThat(result.get(4).getAsString("text"), is("abc"));
			assertThat(result.get(5).getAsString("text"), is("true"));
			assertThat(result.get(6).getAsString("text"), is(nullValue()));
		}
	}

	@Theory
	public void testColumnNumber_int(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 1);
			parser.addColumn("long", "long").set("column_number", 2);
			parser.addColumn("double", "double");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			check_int(result, 0, 123L, 123.4d);
			check_int(result, 1, 456L, 456.7d);
			check_int(result, 2, 123L, 123d);
			check_int(result, 3, 123L, 123.4d);
			check_int(result, 4, 123L, 123.4d);
			check_int(result, 5, 1L, 1d);
			check_int(result, 6, null, null);
		}
	}

	private void check_int(List<OutputRecord> result, int index, Long l, Double d) throws ParseException {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsLong("long"), is(l));
		assertThat(r.getAsDouble("double"), is(d));
	}

	@Theory
	public void testColumnNumber_move(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 1);
			parser.addColumn("long1", "long").set("column_number", 2);
			parser.addColumn("long2", "long").set("column_number", "=");
			parser.addColumn("double1", "double").set("column_number", "+");
			parser.addColumn("double2", "double").set("column_number", "=");
			parser.addColumn("long3", "long").set("column_number", "-");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			check_move(result, 0, 123L, 123.4d);
			check_move(result, 1, 456L, 456.7d);
			check_move(result, 2, 123L, 123d);
			check_move(result, 3, 123L, 123.4d);
			check_move(result, 4, 123L, 123.4d);
			check_move(result, 5, 1L, 1d);
			check_move(result, 6, null, null);
		}
	}

	private void check_move(List<OutputRecord> result, int index, Long l, Double d) throws ParseException {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsLong("long1"), is(l));
		assertThat(r.getAsLong("long2"), is(l));
		assertThat(r.getAsLong("long3"), is(l));
		assertThat(r.getAsDouble("double1"), is(d));
		assertThat(r.getAsDouble("double2"), is(d));
	}

	@Theory
	public void testColumnNumber_move2(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 1);
			parser.addColumn("long1", "long").set("column_number", 2);
			parser.addColumn("string1", "string").set("column_number", "+2");
			parser.addColumn("long2", "long").set("column_number", "=long1");
			parser.addColumn("string2", "string").set("column_number", "=string1");
			parser.addColumn("long3", "long").set("column_number", "-2");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			check_move2(result, 0, 123L, "abc");
			check_move2(result, 1, 456L, "def");
			check_move2(result, 2, 123L, "456");
			check_move2(result, 3, 123L, "abc");
			check_move2(result, 4, 123L, "abc");
			check_move2(result, 5, 1L, "true");
			check_move2(result, 6, null, null);
		}
	}

	private void check_move2(List<OutputRecord> result, int index, Long l, String s) throws ParseException {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsLong("long1"), is(l));
		assertThat(r.getAsLong("long2"), is(l));
		assertThat(r.getAsLong("long3"), is(l));
		assertThat(r.getAsString("string1"), is(s));
		assertThat(r.getAsString("string2"), is(s));
	}

	@Theory
	public void testColumnNumber_moveName(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 1);
			parser.addColumn("long1", "long").set("column_number", 2);
			parser.addColumn("double1", "double").set("column_number", "+long1");
			parser.addColumn("long2", "long").set("column_number", "=long1");
			parser.addColumn("boolean1", "boolean").set("column_number", "-long1");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			check_moveName(result, 0, true, 123L, 123.4d);
			check_moveName(result, 1, false, 456L, 456.7d);
			check_moveName(result, 2, false, 123L, 123d);
			check_moveName(result, 3, true, 123L, 123.4d);
			check_moveName(result, 4, true, 123L, 123.4d);
			check_moveName(result, 5, true, 1L, 1d);
			check_moveName(result, 6, null, null, null);
		}
	}

	private void check_moveName(List<OutputRecord> result, int index, Boolean b, Long l, Double d)
			throws ParseException {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsLong("long1"), is(l));
		assertThat(r.getAsLong("long2"), is(l));
		assertThat(r.getAsDouble("double1"), is(d));
		assertThat(r.getAsBoolean("boolean1"), is(b));
	}
}
