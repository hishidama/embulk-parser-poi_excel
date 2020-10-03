package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
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
public class TestPoiExcelParserPlugin_recordType {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void testRecordType_row(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("record_type", "row");
			parser.set("skip_header_lines", 1);
			parser.addColumn("text", "string").set("cell_column", "D");
			parser.addColumn("text_row", "long").set("cell_column", "D").set("value", "row_number");
			parser.addColumn("text_col", "long").set("cell_column", "D").set("value", "column_number");
			parser.addColumn("text2", "string").set("cell_row", "4");
			parser.addColumn("text2_row", "long").set("cell_row", "4").set("value", "row_number");
			parser.addColumn("text2_col", "long").set("cell_row", "4").set("value", "column_number");
			parser.addColumn("address1", "string").set("cell_address", "B1").set("value", "cell_value");
			parser.addColumn("address2", "string").set("cell_column", "D").set("cell_row", "2");
			parser.addColumn("fix_row", "long").set("cell_address", "B1").set("value", "row_number");
			parser.addColumn("fix_col", "long").set("cell_address", "B1").set("value", "column_number");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			check1(result, 0, "abc");
			check1(result, 1, "def");
			check1(result, 2, "456");
			check1(result, 3, "abc");
			check1(result, 4, "abc");
			check1(result, 5, "true");
			check1(result, 6, null);
		}
	}

	private void check1(List<OutputRecord> result, int index, String text) {
		OutputRecord record = result.get(index);
		// System.out.println(record);
		assertThat(record.getAsString("text"), is(text));
		assertThat(record.getAsLong("text_row"), is((long) index + 2));
		assertThat(record.getAsLong("text_col"), is(4L));
		assertThat(record.getAsString("text2"), is("42283"));
		assertThat(record.getAsLong("text2_row"), is(4L));
		assertThat(record.getAsLong("text2_col"), is(5L));
		assertThat(record.getAsString("address1"), is("long"));
		assertThat(record.getAsString("address2"), is("abc"));
		assertThat(record.getAsLong("fix_row"), is(1L));
		assertThat(record.getAsLong("fix_col"), is(2L));
	}

	@Theory
	public void testRecordType_column0(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("record_type", "column");
			// parser.set("skip_header_lines", 0);
			parser.addColumn("text", "string").set("cell_row", "5");
			parser.addColumn("text_row", "long").set("cell_row", "5").set("value", "row_number");
			parser.addColumn("text_col", "long").set("cell_row", "5").set("value", "column_number");
			parser.addColumn("text2", "string").set("cell_column", "D");
			parser.addColumn("text2_row", "long").set("cell_column", "D").set("value", "row_number");
			parser.addColumn("text2_col", "long").set("cell_column", "D").set("value", "column_number");
			parser.addColumn("address1", "string").set("cell_address", "B1").set("value", "cell_value");
			parser.addColumn("address2", "string").set("cell_column", "D").set("cell_row", "2");
			parser.addColumn("fix_row", "long").set("cell_address", "B1").set("value", "row_number");
			parser.addColumn("fix_col", "long").set("cell_address", "B1").set("value", "column_number");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(7));
			int z = 1;
			check2(result, 0, z, "true");
			check2(result, 1, z, "123");
			check2(result, 2, z, "123.4");
			check2(result, 3, z, "abc");
			check2(result, 4, z, "2015/10/07");
			check2(result, 5, z, null);
			check2(result, 6, z, "CELL_TYPE_STRING");
		}
	}

	@Theory
	public void testRecordType_column1(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("record_type", "column");
			parser.set("skip_header_lines", 1);
			parser.addColumn("text", "string").set("cell_row", "5");
			parser.addColumn("text_row", "long").set("cell_row", "5").set("value", "row_number");
			parser.addColumn("text_col", "long").set("cell_row", "5").set("value", "column_number");
			parser.addColumn("text2", "string").set("cell_column", "D");
			parser.addColumn("text2_row", "long").set("cell_column", "D").set("value", "row_number");
			parser.addColumn("text2_col", "long").set("cell_column", "D").set("value", "column_number");
			parser.addColumn("address1", "string").set("cell_address", "B1").set("value", "cell_value");
			parser.addColumn("address2", "string").set("cell_column", "D").set("cell_row", "2");
			parser.addColumn("fix_row", "long").set("cell_address", "B1").set("value", "row_number");
			parser.addColumn("fix_col", "long").set("cell_address", "B1").set("value", "column_number");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(6));
			int z = 2;
			check2(result, 0, z, "123");
			check2(result, 1, z, "123.4");
			check2(result, 2, z, "abc");
			check2(result, 3, z, "2015/10/07");
			check2(result, 4, z, null);
			check2(result, 5, z, "CELL_TYPE_STRING");
		}
	}

	private void check2(List<OutputRecord> result, int index, int z, String text) {
		OutputRecord record = result.get(index);
		// System.out.println(record);
		assertThat(record.getAsString("text"), is(text));
		assertThat(record.getAsLong("text_row"), is(5L));
		assertThat(record.getAsLong("text_col"), is((long) index + z));
		assertThat(record.getAsString("text2"), is("abc"));
		assertThat(record.getAsLong("text2_row"), is(6L));
		assertThat(record.getAsLong("text2_col"), is(4L));
		assertThat(record.getAsString("address1"), is("long"));
		assertThat(record.getAsString("address2"), is("abc"));
		assertThat(record.getAsLong("fix_row"), is(1L));
		assertThat(record.getAsLong("fix_col"), is(2L));
	}

	@Theory
	public void testRecordType_sheet(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "style");
			parser.set("record_type", "sheet");
			parser.addColumn("text", "string");
			parser.addColumn("text_row", "long").set("cell_row", "5").set("value", "row_number");
			parser.addColumn("text_col", "long").set("cell_column", "6").set("value", "column_number");
			parser.addColumn("address1", "string").set("cell_address", "B1").set("value", "cell_value");
			parser.addColumn("address2", "string").set("cell_column", "A").set("cell_row", "4");
			parser.addColumn("fix_row", "long").set("cell_address", "B1").set("value", "row_number");
			parser.addColumn("fix_col", "long").set("cell_address", "B1").set("value", "column_number");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(1));
			check3(result, 0, "red");
		}
	}

	private void check3(List<OutputRecord> result, int index, String text) {
		OutputRecord record = result.get(index);
		// System.out.println(record);
		assertThat(record.getAsString("text"), is(text));
		assertThat(record.getAsLong("text_row"), is(5L));
		assertThat(record.getAsLong("text_col"), is(6L));
		assertThat(record.getAsString("address1"), is("top"));
		assertThat(record.getAsString("address2"), is("white"));
		assertThat(record.getAsLong("fix_row"), is(1L));
		assertThat(record.getAsLong("fix_col"), is(2L));
	}
}
