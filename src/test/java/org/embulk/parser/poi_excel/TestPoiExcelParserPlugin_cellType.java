package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.MatcherAssert.assertThat;

import java.net.URL;
import java.text.ParseException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.CellType;
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
			check1(result, 0, CellType.NUMERIC, "NUMERIC");
			check1(result, 1, CellType.STRING, "STRING");
			check1(result, 2, CellType.FORMULA, "FORMULA");
			check1(result, 3, CellType.BOOLEAN, "BOOLEAN");
			check1(result, 4, CellType.FORMULA, "FORMULA");
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
			check1(result, 0, CellType.NUMERIC, "NUMERIC");
			check1(result, 1, CellType.STRING, "STRING");
			check1(result, 2, CellType.BOOLEAN, "BOOLEAN");
			check1(result, 3, CellType.BOOLEAN, "BOOLEAN");
			check1(result, 4, CellType.ERROR, "ERROR");
		}
	}

	@SuppressWarnings("deprecation")
	private void check1(List<OutputRecord> result, int index, CellType cellType, String s) throws ParseException {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsLong("long"), is((long) cellType.getCode()));
		assertThat(r.getAsString("string"), is(s));
	}
}
