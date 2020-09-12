package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.junit.Assert.fail;

import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.List;

import org.apache.poi.ss.usermodel.FormulaError;
import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.embulk.spi.time.Timestamp;
import org.junit.experimental.theories.DataPoints;
import org.junit.experimental.theories.Theories;
import org.junit.experimental.theories.Theory;
import org.junit.runner.RunWith;

@RunWith(Theories.class)
public class TestPoiExcelParserPlugin_cellError {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void testCellError_default(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 7);
			parser.addColumn("b", "boolean").set("column_number", "A");
			parser.addColumn("l", "long").set("column_number", "A");
			parser.addColumn("d", "double").set("column_number", "A");
			parser.addColumn("s", "string").set("column_number", "A");
			parser.addColumn("t", "timestamp").set("column_number", "A");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(1));
			OutputRecord r = result.get(0);
			assertThat(r.getAsBoolean("b"), is(nullValue()));
			assertThat(r.getAsLong("l"), is(nullValue()));
			assertThat(r.getAsDouble("d"), is(nullValue()));
			assertThat(r.getAsString("s"), is(nullValue()));
			assertThat(r.getAsString("t"), is(nullValue()));
		}
	}

	@Theory
	public void testCellError_code(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 7);
			parser.set("on_cell_error", "error_code");
			parser.addColumn("b", "boolean").set("column_number", "A");
			parser.addColumn("l", "long").set("column_number", "A");
			parser.addColumn("d", "double").set("column_number", "A");
			parser.addColumn("s", "string").set("column_number", "A");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			OutputRecord r = result.get(0);
			assertThat(r.getAsBoolean("b"), is(nullValue()));
			assertThat(r.getAsLong("l"), is((long) FormulaError.DIV0.getCode()));
			assertThat(r.getAsDouble("d"), is((double) FormulaError.DIV0.getCode()));
			assertThat(r.getAsString("s"), is("#DIV/0!"));
		}
	}

	@Theory
	public void testCellError_null(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 7);
			parser.set("on_cell_error", "constant");
			parser.addColumn("b", "boolean").set("column_number", "A");
			parser.addColumn("l", "long").set("column_number", "A");
			parser.addColumn("d", "double").set("column_number", "A");
			parser.addColumn("s", "string").set("column_number", "A");
			parser.addColumn("t", "timestamp").set("column_number", "A");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(1));
			OutputRecord r = result.get(0);
			assertThat(r.getAsBoolean("b"), is(nullValue()));
			assertThat(r.getAsLong("l"), is(nullValue()));
			assertThat(r.getAsDouble("d"), is(nullValue()));
			assertThat(r.getAsString("s"), is(nullValue()));
			assertThat(r.getAsString("t"), is(nullValue()));
		}
	}

	@Theory
	public void testCellError_empty(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 7);
			parser.set("on_cell_error", "constant.zzz");
			parser.addColumn("s1", "string").set("column_number", "A");
			parser.addColumn("s2", "string").set("column_number", "A").set("on_cell_error", "constant");
			parser.addColumn("s3", "string").set("column_number", "A").set("on_cell_error", "constant.");
			parser.addColumn("s4", "string").set("column_number", "A").set("on_cell_error", "constant. ");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(1));
			OutputRecord r = result.get(0);
			assertThat(r.getAsString("s1"), is("zzz"));
			assertThat(r.getAsString("s2"), is(nullValue()));
			assertThat(r.getAsString("s3"), is(""));
			assertThat(r.getAsString("s4"), is(" "));
		}
	}

	@Theory
	public void testCellError_constant(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 7);
			parser.set("on_cell_error", "constant.0");
			parser.addColumn("b", "boolean").set("column_number", "A");
			parser.addColumn("l", "long").set("column_number", "A");
			parser.addColumn("d", "double").set("column_number", "A");
			parser.addColumn("s", "string").set("column_number", "A");
			parser.addColumn("t", "timestamp").set("column_number", "A").set("format", "%Y/%m/%d")
					.set("on_cell_error", "constant.2000/1/1");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			OutputRecord r = result.get(0);
			assertThat(r.getAsBoolean("b"), is(false));
			assertThat(r.getAsLong("l"), is(0L));
			assertThat(r.getAsDouble("d"), is(0d));
			assertThat(r.getAsString("s"), is("0"));
			assertThat(r.getAsTimestamp("t"),
					is(Timestamp.ofEpochMilli(new SimpleDateFormat("yyyy/MM/dd z").parse("2000/01/01 UTC").getTime())));
		}
	}

	@Theory
	public void testCellError_exception(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 7);
			parser.set("on_cell_error", "exception");
			parser.addColumn("b", "boolean").set("column_number", "A");
			parser.addColumn("l", "long").set("column_number", "A");
			parser.addColumn("d", "double").set("column_number", "A");
			parser.addColumn("s", "string").set("column_number", "A");
			parser.addColumn("t", "timestamp").set("column_number", "A");

			URL inFile = getClass().getResource(excelFile);
			try {
				tester.runParser(inFile, parser);
			} catch (Exception e) {
				Throwable c1 = e.getCause();
				assertThat(c1.getMessage().contains("error at Column"), is(true));
				Throwable c2 = c1.getCause();
				assertThat(c2.getMessage().contains("encount cell error"), is(true));
				return; // success
			}
			fail("must throw Exception");
		}
	}
}
