package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

import java.net.URL;
import java.text.ParseException;
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
}
