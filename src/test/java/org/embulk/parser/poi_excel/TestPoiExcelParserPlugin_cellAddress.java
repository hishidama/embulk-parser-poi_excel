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
public class TestPoiExcelParserPlugin_cellAddress {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void testCellAddress(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "test1");
			parser.set("skip_header_lines", 1);
			parser.addColumn("text", "string").set("column_number", "D");
			parser.addColumn("fix_value", "string").set("cell_address", "B1").set("value", "cell_value");
			parser.addColumn("fix_sheet", "string").set("cell_address", "B1").set("value", "sheet_name");
			parser.addColumn("fix_row", "long").set("cell_address", "B1").set("value", "row_number");
			parser.addColumn("fix_col", "long").set("cell_address", "B1").set("value", "column_number");
			parser.addColumn("other_sheet_value", "string").set("cell_address", "style!B5").set("value", "cell_value");
			parser.addColumn("other_sheet_name", "string").set("cell_address", "style!B5").set("value", "sheet_name");
			parser.addColumn("other_sheet_row", "long").set("cell_address", "style!B5").set("value", "row_number");
			parser.addColumn("other_sheet_col", "string").set("cell_address", "style!B5").set("value", "column_number");

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
		assertThat(record.getAsString("fix_value"), is("long"));
		assertThat(record.getAsString("fix_sheet"), is("test1"));
		assertThat(record.getAsLong("fix_row"), is(1L));
		assertThat(record.getAsLong("fix_col"), is(2L));
		assertThat(record.getAsString("other_sheet_value"), is("bottom"));
		assertThat(record.getAsString("other_sheet_name"), is("style"));
		assertThat(record.getAsLong("other_sheet_row"), is(5L));
		assertThat(record.getAsString("other_sheet_col"), is("B"));
	}
}
