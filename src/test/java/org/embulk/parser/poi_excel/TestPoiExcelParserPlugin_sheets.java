package org.embulk.parser.poi_excel;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

import java.net.URL;
import java.text.ParseException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.embulk.parser.EmbulkPluginTester;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.parser.EmbulkTestParserConfig;
import org.junit.Test;

public class TestPoiExcelParserPlugin_sheets {

	@Test
	public void testSheets() throws ParseException {
		try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
			tester.addParserPlugin(PoiExcelParserPlugin.TYPE, PoiExcelParserPlugin.class);

			EmbulkTestParserConfig parser = tester.newParserConfig(PoiExcelParserPlugin.TYPE);
			parser.set("sheets", Arrays.asList("test1", "formula_replace", "style"));
			parser.addColumn("text", "string");
			parser.addColumn("number", "long");

			Map<String, Object> sheetOptions = new HashMap<>();
			{
				Map<String, Object> sheet = new HashMap<>();
				sheet.put("skip_header_lines", "5");
				Map<String, Object> columns = new HashMap<>();
				columns.put("text", newMap("column_number", "D"));
				columns.put("number", newMap("column_number", "B"));
				sheet.put("columns", columns);
				sheetOptions.put("test1", sheet);
			}
			{
				Map<String, Object> sheet = new HashMap<>();
				Map<String, Object> columns = new HashMap<>();
				columns.put("number", newMap("value", "constant.0"));
				sheet.put("columns", columns);
				sheetOptions.put("formula_replace", sheet);
			}
			{
				Map<String, Object> sheet = new HashMap<>();
				sheet.put("skip_header_lines", "2");
				Map<String, Object> columns = new HashMap<>();
				columns.put("text", newMap("column_number", "B"));
				columns.put("number", newMap("value", "constant.-1"));
				sheet.put("columns", columns);
				sheetOptions.put("style", sheet);
			}
			parser.set("sheet_options", sheetOptions);

			URL inFile = getClass().getResource("test1.xls");
			List<OutputRecord> result = tester.runParser(inFile, parser);

			assertThat(result.size(), is(8));
			check1(result, 0, "abc", 123L);
			check1(result, 1, "true", 1L);
			check1(result, 2, null, null);
			check1(result, 3, "boolean", 0L);
			check1(result, 4, "test2-b1", 0L);
			check1(result, 5, "left", -1L);
			check1(result, 6, "right", -1L);
			check1(result, 7, "bottom", -1L);
		}
	}

	private Map<String, Object> newMap(String key, Object value) {
		Map<String, Object> map = new HashMap<>();
		map.put(key, value);
		return map;
	}

	private void check1(List<OutputRecord> result, int index, String text, Long number) {
		OutputRecord record = result.get(index);
		// System.out.println(record);
		assertThat(record.getAsString("text"), is(text));
		assertThat(record.getAsLong("number"), is(number));
	}
}
