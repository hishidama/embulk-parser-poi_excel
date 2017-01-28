package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

import java.net.URL;
import java.text.ParseException;
import java.util.Arrays;
import java.util.List;

import org.embulk.config.ConfigSource;
import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.junit.experimental.theories.DataPoints;
import org.junit.experimental.theories.Theories;
import org.junit.experimental.theories.Theory;
import org.junit.runner.RunWith;

@RunWith(Theories.class)
public class TestPoiExcelParserPlugin_formula {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void testForumlaHandlingCashedValue(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "formula_replace");

			parser.addColumn("text", "string").set("formula_handling", "cashed_value");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(2));
			assertThat(result.get(0).getAsString("text"), is("boolean"));
			assertThat(result.get(1).getAsString("text"), is("test2-b1"));
		}
	}

	@Theory
	public void testForumlaHandlingEvaluate(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "formula_replace");

			parser.addColumn("text", "string").set("formula_handling", "evaluate");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(2));
			assertThat(result.get(0).getAsString("text"), is("boolean"));
			assertThat(result.get(1).getAsString("text"), is("test2-b1"));
		}
	}

	@Theory
	public void testForumlaReplace(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "formula_replace");

			ConfigSource replace0 = tester.newConfigSource();
			replace0.set("regex", "test1");
			replace0.set("to", "merged_cell");
			ConfigSource replace1 = tester.newConfigSource();
			replace1.set("regex", "B1");
			replace1.set("to", "B${row}");
			parser.set("formula_replace", Arrays.asList(replace0, replace1));

			parser.addColumn("text", "string");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(2));
			assertThat(result.get(0).getAsString("text"), is("test3-a1"));
			assertThat(result.get(1).getAsString("text"), is("test2-b2"));
		}
	}
}
