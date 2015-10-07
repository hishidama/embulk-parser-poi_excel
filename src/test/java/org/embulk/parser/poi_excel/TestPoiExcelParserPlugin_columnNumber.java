package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

import java.net.URL;
import java.text.ParseException;
import java.util.List;

import org.apache.poi.ss.usermodel.FormulaError;
import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.junit.Test;

public class TestPoiExcelParserPlugin_columnNumber {

	@Test
	public void testColumnNumber_string() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 1);
			parser.set("cell_error_null", false);
			parser.addColumn("text", "string").set("column_number", "D");

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			assertThat(result.get(0).getAsString("text"), is("abc"));
			assertThat(result.get(1).getAsString("text"), is("def"));
			assertThat(result.get(2).getAsString("text"), is("456"));
			assertThat(result.get(3).getAsString("text"), is("abc"));
			assertThat(result.get(4).getAsString("text"), is("abc"));
			assertThat(result.get(5).getAsString("text"), is("true"));
			assertThat(result.get(6).getAsString("text"), is("#DIV/0!"));
		}
	}

	@Test
	public void testColumnNumber_int() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 1);
			parser.set("cell_error_null", false);
			parser.addColumn("long", "long").set("column_number", 2);
			parser.addColumn("double", "double");

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			check_int(result, 0, 123L, 123.4d);
			check_int(result, 1, 456L, 456.7d);
			check_int(result, 2, 123L, 123d);
			check_int(result, 3, 123L, 123.4d);
			check_int(result, 4, 123L, 123.4d);
			check_int(result, 5, 1L, 1d);
			check_int(result, 6, (long) FormulaError.DIV0.getCode(), (double) FormulaError.DIV0.getCode());
		}
	}

	private void check_int(List<OutputRecord> result, int index, Long l, Double d) throws ParseException {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsLong("long"), is(l));
		assertThat(r.getAsDouble("double"), is(d));
	}

	@Test
	public void testColumnNumber_move() throws ParseException {
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

			URL inFile = getClass().getResource("test1.xls");
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
}
