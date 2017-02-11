package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

import java.net.URL;
import java.text.ParseException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.junit.experimental.theories.DataPoints;
import org.junit.experimental.theories.Theories;
import org.junit.experimental.theories.Theory;
import org.junit.runner.RunWith;

@RunWith(Theories.class)
public class TestPoiExcelParserPlugin_cellType {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void testCellType(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheets", Arrays.asList("test1"));
			parser.set("skip_header_lines", 3);
			parser.addColumn("long", "long").set("column_number", "A").set("value", "cell_type");
			parser.addColumn("string", "string").set("column_number", "A").set("value", "cell_type");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(5));
			check1(result, 0, Cell.CELL_TYPE_NUMERIC, "NUMERIC");
			check1(result, 1, Cell.CELL_TYPE_STRING, "STRING");
			check1(result, 2, Cell.CELL_TYPE_FORMULA, "FORMULA");
			check1(result, 3, Cell.CELL_TYPE_BOOLEAN, "BOOLEAN");
			check1(result, 4, Cell.CELL_TYPE_FORMULA, "FORMULA");
		}
	}

	@Theory
	public void testCellCachedType(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheets", Arrays.asList("test1"));
			parser.set("skip_header_lines", 3);
			parser.addColumn("long", "long").set("column_number", "A").set("value", "cell_cached_type");
			parser.addColumn("string", "string").set("column_number", "A").set("value", "cell_cached_type");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(5));
			check1(result, 0, Cell.CELL_TYPE_NUMERIC, "NUMERIC");
			check1(result, 1, Cell.CELL_TYPE_STRING, "STRING");
			check1(result, 2, Cell.CELL_TYPE_BOOLEAN, "BOOLEAN");
			check1(result, 3, Cell.CELL_TYPE_BOOLEAN, "BOOLEAN");
			check1(result, 4, Cell.CELL_TYPE_ERROR, "ERROR");
		}
	}

	private void check1(List<OutputRecord> result, int index, long l, String s) throws ParseException {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsLong("long"), is(l));
		assertThat(r.getAsString("string"), is(s));
	}
}
