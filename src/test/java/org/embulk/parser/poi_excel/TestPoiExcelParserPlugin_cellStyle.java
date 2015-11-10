package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;
import static org.junit.Assert.fail;

import java.net.URL;
import java.text.ParseException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.junit.experimental.theories.DataPoints;
import org.junit.experimental.theories.Theories;
import org.junit.experimental.theories.Theory;
import org.junit.runner.RunWith;

@RunWith(Theories.class)
public class TestPoiExcelParserPlugin_cellStyle {

	@DataPoints
	public static String[] FILES = { "test1.xls", "test2.xlsx" };

	@Theory
	public void testStyle_key(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "style");
			parser.addColumn("color-text", "string");
			parser.addColumn("color", "string").set("value", "cell_style.fill_foreground_color");
			parser.addColumn("border-text", "string");
			parser.addColumn("border-top", "long").set("value", "cell_style.border_top");
			parser.addColumn("border-bottom", "long").set("value", "cell_style.border_bottom");
			parser.addColumn("border-left", "long").set("value", "cell_style.border_left");
			parser.addColumn("border-right", "long").set("value", "cell_style.border_right");
			parser.addColumn("border-all", "long").set("value", "cell_style.border");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(5));
			check1(result, 0, "red", 255, 0, 0, "top", CellStyle.BORDER_THIN, 0, 0, 0);
			check1(result, 1, "green", 0, 128, 0, null, 0, 0, 0, 0);
			check1(result, 2, "blue", 0, 0, 255, "left", 0, 0, CellStyle.BORDER_THIN, 0);
			check1(result, 3, "white", 255, 255, 255, "right", 0, 0, 0, CellStyle.BORDER_THIN);
			check1(result, 4, "black", 0, 0, 0, "bottom", 0, CellStyle.BORDER_MEDIUM, 0, 0);
		}
	}

	private void check1(List<OutputRecord> result, int index, String colorText, int r, int g, int b, String borderText,
			long top, long bottom, long left, long right) {
		OutputRecord record = result.get(index);
		// System.out.println(record);
		assertThat(record.getAsString("color-text"), is(colorText));
		assertThat(record.getAsString("color"), is(String.format("%02x%02x%02x", r, g, b)));
		assertThat(record.getAsString("border-text"), is(borderText));
		assertThat(record.getAsLong("border-top"), is(top));
		assertThat(record.getAsLong("border-bottom"), is(bottom));
		assertThat(record.getAsLong("border-left"), is(left));
		assertThat(record.getAsLong("border-right"), is(right));
		assertThat(record.getAsLong("border-all"), is(top << 24 | bottom << 16 | left << 8 | right));
	}

	@Theory
	public void testStyle_all(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "style");
			parser.addColumn("color-text", "string");
			parser.addColumn("color-style", "string").set("column_number", "A").set("value", "cell_style");
			parser.addColumn("border-style", "string").set("column_number", "B").set("value", "cell_style");

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(5));
			check2(result, 0, "red", 255, 0, 0, "top", CellStyle.BORDER_THIN, 0, 0, 0);
			check2(result, 1, "green", 0, 128, 0, null, 0, 0, 0, 0);
			check2(result, 2, "blue", 0, 0, 255, "left", 0, 0, CellStyle.BORDER_THIN, 0);
			check2(result, 3, "white", 255, 255, 255, "right", 0, 0, 0, CellStyle.BORDER_THIN);
			check2(result, 4, "black", 0, 0, 0, "bottom", 0, CellStyle.BORDER_MEDIUM, 0, 0);
		}
	}

	private void check2(List<OutputRecord> result, int index, String colorText, int r, int g, int b, String borderText,
			long top, long bottom, long left, long right) {
		OutputRecord record = result.get(index);
		// System.out.println(record);
		assertThat(record.getAsString("color-text"), is(colorText));
		String color = record.getAsString("color-style");
		if (!color.contains(String.format("\"fill_foreground_color\":\"%02x%02x%02x\"", r, g, b))) {
			fail(color);
		}
		String border = record.getAsString("border-style");
		if (!border.contains(String.format("\"border_top\":%d", top))) {
			fail(border);
		}
		if (!border.contains(String.format("\"border_bottom\":%d", bottom))) {
			fail(border);
		}
		if (!border.contains(String.format("\"border_left\":%d", left))) {
			fail(border);
		}
		if (!border.contains(String.format("\"border_right\":%d", right))) {
			fail(border);
		}
	}

	@Theory
	public void testStyle_keys(String excelFile) throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheet", "style");
			parser.addColumn("color-text", "string");
			parser.addColumn("color-style", "string").set("column_number", "A").set("value", "cell_style")
					.set("cell_style_name", Arrays.asList("fill_foreground_color"));
			parser.addColumn("border-style", "string").set("column_number", "B").set("value", "cell_style")
					.set("attribute_name", Arrays.asList("border_top", "border_bottom", "border_left", "border_right"));

			URL inFile = getClass().getResource(excelFile);
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(5));
			check2(result, 0, "red", 255, 0, 0, "top", CellStyle.BORDER_THIN, 0, 0, 0);
			check2(result, 1, "green", 0, 128, 0, null, 0, 0, 0, 0);
			check2(result, 2, "blue", 0, 0, 255, "left", 0, 0, CellStyle.BORDER_THIN, 0);
			check2(result, 3, "white", 255, 255, 255, "right", 0, 0, 0, CellStyle.BORDER_THIN);
			check2(result, 4, "black", 0, 0, 0, "bottom", 0, CellStyle.BORDER_MEDIUM, 0, 0);
		}
	}
}
