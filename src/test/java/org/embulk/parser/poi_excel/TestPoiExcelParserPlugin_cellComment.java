package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.junit.Assert.assertThat;
import static org.junit.Assert.fail;

import java.net.URL;
import java.text.ParseException;
import java.util.Arrays;
import java.util.List;

import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.junit.Test;

public class TestPoiExcelParserPlugin_cellComment {

	@Test
	public void testComment_key() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "comment");
			parser.addColumn("author", "string").set("value", "cell_comment.author");
			parser.addColumn("comment", "string").set("value", "cell_comment.string");

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(2));
			check1(result, 0, "hishidama", "hishidama:\nmy comment");
			check1(result, 1, null, null);
		}
	}

	private void check1(List<OutputRecord> result, int index, String author, String comment) {
		OutputRecord record = result.get(index);
		// System.out.println(record);
		assertThat(record.getAsString("comment"), is(comment));
		assertThat(record.getAsString("author"), is(author));
	}

	@Test
	public void testComment_all() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "comment");
			parser.addColumn("comment", "string").set("value", "cell_comment");

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(2));
			check2(result, 0, "hishidama", "hishidama:\\nmy comment");
			check2(result, 1, null, null);
		}
	}

	private void check2(List<OutputRecord> result, int index, String author, String comment) {
		OutputRecord record = result.get(index);
		// System.out.println(record);
		String s = record.getAsString("comment");
		if (author == null && comment == null) {
			assertThat(s, is(nullValue()));
			return;
		}

		if (!s.contains(String.format("\"author\":\"%s\"", author))) {
			fail(s);
		}
		if (!s.contains(String.format("\"string\":\"%s\"", comment))) {
			fail(s);
		}
	}

	@Test
	public void testComment_keys() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "comment");
			parser.addColumn("comment", "string").set("value", "cell_comment")
					.set("attribute_name", Arrays.asList("author", "string"));

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(2));
			check2(result, 0, "hishidama", "hishidama:\\nmy comment");
			check2(result, 1, null, null);
		}
	}
}
