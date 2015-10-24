package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.junit.Assert.assertThat;

import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.List;

import org.apache.poi.ss.usermodel.FormulaError;
import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.embulk.spi.time.Timestamp;
import org.junit.Test;

public class TestPoiExcelParserPlugin_cellError {

	@Test
	public void testCellError_default() throws Exception {
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

			URL inFile = getClass().getResource("test1.xls");
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

	@Test
	public void testCellError_code() throws Exception {
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
			parser.addColumn("t", "timestamp").set("column_number", "A");

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			OutputRecord r = result.get(0);
			assertThat(r.getAsBoolean("b"), is(nullValue()));
			assertThat(r.getAsLong("l"), is((long) FormulaError.DIV0.getCode()));
			assertThat(r.getAsDouble("d"), is((double) FormulaError.DIV0.getCode()));
			assertThat(r.getAsString("s"), is("#DIV/0!"));
			assertThat(r.getAsString("t"), is(nullValue()));
		}
	}

	@Test
	public void testCellError_null() throws Exception {
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

			URL inFile = getClass().getResource("test1.xls");
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

	@Test
	public void testCellError_empty() throws Exception {
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

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(1));
			OutputRecord r = result.get(0);
			assertThat(r.getAsString("s1"), is("zzz"));
			assertThat(r.getAsString("s2"), is(nullValue()));
			assertThat(r.getAsString("s3"), is(""));
			assertThat(r.getAsString("s4"), is(" "));
		}
	}

	@Test
	public void testCellError_constant() throws Exception {
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

			URL inFile = getClass().getResource("test1.xls");
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
}
