package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
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
public class TestPoiExcelParserPlugin_constant {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void testConstant(String excelFile) throws Exception {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "style");
			parser.addColumn("const-s", "string").set("value", "constant.zzz");
			parser.addColumn("const-n", "long").set("value", "constant.-1");
			parser.addColumn("space", "string").set("value", "constant. ");
			parser.addColumn("empty", "string").set("value", "constant.");
			parser.addColumn("null", "string").set("value", "constant");
			parser.addColumn("cell", "string");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(5));
			check(result, 0, "red");
			check(result, 1, "green");
			check(result, 2, "blue");
			check(result, 3, "white");
			check(result, 4, "black");
		}
	}

	private void check(List<OutputRecord> result, int index, String s) throws ParseException {
		OutputRecord r = result.get(index);
		// System.out.println(r);
		assertThat(r.getAsString("const-s"), is("zzz"));
		assertThat(r.getAsLong("const-n"), is(-1L));
		assertThat(r.getAsString("space"), is(" "));
		assertThat(r.getAsString("empty"), is(""));
		assertThat(r.getAsString("null"), is(nullValue()));
		assertThat(r.getAsString("cell"), is(s));
	}
}
