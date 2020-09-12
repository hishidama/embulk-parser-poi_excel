package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.junit.Assert.fail;

import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.List;

import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.embulk.spi.time.Timestamp;
import org.junit.experimental.theories.DataPoints;
import org.junit.experimental.theories.Theories;
import org.junit.experimental.theories.Theory;
import org.junit.runner.RunWith;

@RunWith(Theories.class)
public class TestPoiExcelParserPlugin_convertError {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void testConvertError_default(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "style");
			parser.addColumn("t", "timestamp").set("column_number", "A");

			URL inFile = getClass().getResource(excelFile);
			try {
				tester.runParser(inFile, parser);
			} catch (Exception e) {
				Throwable c1 = e.getCause();
				assertThat(c1.getMessage().contains("error at Column"), is(true));
				Throwable c2 = c1.getCause();
				assertThat(c2.getMessage().contains("convert error"), is(true));
				return; // success
			}
			fail("must throw Exception");
		}
	}

	@Theory
	public void testConvertError_exception(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "style");
			parser.set("on_convert_error", "exception");
			parser.addColumn("t", "timestamp").set("column_number", "A");

			URL inFile = getClass().getResource(excelFile);
			try {
				tester.runParser(inFile, parser);
			} catch (Exception e) {
				Throwable c1 = e.getCause();
				assertThat(c1.getMessage().contains("error at Column"), is(true));
				Throwable c2 = c1.getCause();
				assertThat(c2.getMessage().contains("convert error"), is(true));
				return; // success
			}
			fail("must throw Exception");
		}
	}

	@Theory
	public void testConvertError_null(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "comment");
			parser.set("on_convert_error", "constant");
			parser.addColumn("t", "timestamp").set("column_number", "A");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(2));
			assertThat(result.get(0).getAsTimestamp("t"), is(nullValue()));
			assertThat(result.get(1).getAsTimestamp("t"), is(nullValue()));
		}
	}

	@Theory
	public void testConvertError_constant(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "comment");
			parser.set("on_convert_error", "constant.0");
			parser.addColumn("b", "boolean").set("column_number", "A");
			parser.addColumn("l", "long").set("column_number", "A");
			parser.addColumn("d", "double").set("column_number", "A");
			parser.addColumn("t", "timestamp").set("column_number", "A").set("format", "%Y/%m/%d")
					.set("on_convert_error", "constant.2000/1/1");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(2));
			for (OutputRecord r : result) {
				assertThat(r.getAsBoolean("b"), is(false));
				assertThat(r.getAsLong("l"), is(0L));
				assertThat(r.getAsDouble("d"), is(0d));
				assertThat(r.getAsTimestamp("t"), is(Timestamp.ofEpochMilli(new SimpleDateFormat("yyyy/MM/dd z").parse(
						"2000/01/01 UTC").getTime())));
			}
		}
	}
}
