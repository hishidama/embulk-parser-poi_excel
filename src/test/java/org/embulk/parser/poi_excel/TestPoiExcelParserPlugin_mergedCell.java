package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

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
public class TestPoiExcelParserPlugin_mergedCell {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void testSearchMergedCell_default(String excelFile) throws ParseException {
		test(excelFile, null, true);
	}

	@Theory
	public void testSearchMergedCell_true(String excelFile) throws ParseException {
		// compatibility ver 0.1.7
		test(excelFile, true, true);
	}

	@Theory
	public void testSearchMergedCell_false(String excelFile) throws ParseException {
		// compatibility ver 0.1.7
		test(excelFile, false, false);
	}

	@Theory
	public void testSearchMergedCell_none(String excelFile) throws ParseException {
		test(excelFile, "none", false);
	}

	@Theory
	public void testSearchMergedCell_linear(String excelFile) throws ParseException {
		test(excelFile, "linear_search", true);
	}

	@Theory
	public void testSearchMergedCell_tree(String excelFile) throws ParseException {
		test(excelFile, "tree_search", true);
	}

	private void test(String excelFile, Object arg, boolean search) {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "merged_cell");
			if (arg != null) {
				parser.set("search_merged_cell", arg);
			}
			parser.addColumn("a", "string");
			parser.addColumn("b", "string");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(4));
			if (search) {
				check6(result, 0, "test3-a1", "test3-a1");
			} else {
				check6(result, 0, "test3-a1", null);
			}
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
