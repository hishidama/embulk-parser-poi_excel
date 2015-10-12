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

public class TestPoiExcelParserPlugin_cellFont {

	@Test
	public void testFont_key() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "style");
			parser.addColumn("color-text", "string");
			parser.addColumn("font-color", "long").set("column_number", "C").set("value_type", "cell_font.color");
			parser.addColumn("font-bold", "boolean").set("value_type", "cell_font.bold");

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(5));
			check1(result, 0, "red", null, false);
			check1(result, 1, "green", 0xff0000L, true);
			check1(result, 2, "blue", null, null);
			check1(result, 3, "white", null, null);
			check1(result, 4, "black", null, null);
		}
	}

	private void check1(List<OutputRecord> result, int index, String colorText, Long fontColor, Boolean fontBold) {
		OutputRecord record = result.get(index);
		// System.out.println(record);
		assertThat(record.getAsString("color-text"), is(colorText));
		assertThat(record.getAsLong("font-color"), is(fontColor));
		assertThat(record.getAsBoolean("font-bold"), is(fontBold));
	}

	@Test
	public void testFont_all() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "style");
			parser.addColumn("color-text", "string");
			parser.addColumn("color-font", "string").set("column_number", "C").set("value_type", "cell_font");

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(5));
			check2(result, 0, "red", null, false);
			check2(result, 1, "green", 0xff0000L, true);
			check2(result, 2, "blue", null, null);
			check2(result, 3, "white", null, null);
			check2(result, 4, "black", null, null);
		}
	}

	private void check2(List<OutputRecord> result, int index, String colorText, Long fontColor, Boolean fontBold) {
		OutputRecord record = result.get(index);
		// System.out.println(record);
		assertThat(record.getAsString("color-text"), is(colorText));
		String font = record.getAsString("color-font");
		if (fontColor == null && fontBold == null) {
			assertThat(font, is(nullValue()));
			return;
		}

		if (fontColor == null) {
			if (!font.contains("\"color\":null")) {
				fail(font);
			}
		} else {
			if (!font.contains(String.format("\"color\":\"%06x\"", fontColor))) {
				fail(font);
			}
		}
		if (fontBold == null) {
			if (!font.contains("\"bold\":null")) {
				fail(font);
			}
		} else {
			if (!font.contains(String.format("\"bold\":%b", fontBold))) {
				fail(font);
			}
		}
	}

	@Test
	public void testFont_keys() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "style");
			parser.addColumn("color-text", "string");
			parser.addColumn("color-font", "string").set("column_number", "C").set("value_type", "cell_font")
					.set("attribute_name", Arrays.asList("color", "bold"));

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(5));
			check2(result, 0, "red", null, false);
			check2(result, 1, "green", 0xff0000L, true);
			check2(result, 2, "blue", null, null);
			check2(result, 3, "white", null, null);
			check2(result, 4, "black", null, null);
		}
	}
}
