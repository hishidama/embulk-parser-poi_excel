package org.embulk.parser.poi_excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.embulk.config.Config;
import org.embulk.config.ConfigDefault;
import org.embulk.config.ConfigException;
import org.embulk.config.ConfigSource;
import org.embulk.config.Task;
import org.embulk.config.TaskSource;
import org.embulk.parser.poi_excel.bean.PoiExcelSheetBean;
import org.embulk.parser.poi_excel.bean.record.PoiExcelRecord;
import org.embulk.parser.poi_excel.visitor.PoiExcelColumnVisitor;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorFactory;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorValue;
import org.embulk.spi.Exec;
import org.embulk.spi.FileInput;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.PageOutput;
import org.embulk.spi.ParserPlugin;
import org.embulk.spi.Schema;
import org.embulk.spi.SchemaConfig;
import org.embulk.spi.time.TimestampParser;
import org.embulk.spi.util.FileInputInputStream;
import org.slf4j.Logger;

import com.google.common.base.Optional;

public class PoiExcelParserPlugin implements ParserPlugin {
	private final Logger log = Exec.getLogger(getClass());

	public static final String TYPE = "poi_excel";

	public interface PluginTask extends Task, TimestampParser.Task, SheetCommonOptionTask {
		@Config("sheet")
		@ConfigDefault("null")
		public Optional<String> getSheet();

		@Config("sheets")
		@ConfigDefault("[]")
		public List<String> getSheets();

		@Config("ignore_sheet_not_found")
		@ConfigDefault("false")
		public boolean getIgnoreSheetNotFound();

		@Config("sheet_options")
		@ConfigDefault("{}")
		public Map<String, SheetOptionTask> getSheetOptions();

		@Config("columns")
		public SchemaConfig getColumns();

		@Config("flush_count")
		@ConfigDefault("100")
		public int getFlushCount();
	}

	public interface SheetCommonOptionTask extends Task, ColumnCommonOptionTask {

		@Config("record_type")
		@ConfigDefault("null")
		public Optional<String> getRecordType();

		@Config("skip_header_lines")
		@ConfigDefault("null")
		public Optional<Integer> getSkipHeaderLines();
	}

	public interface SheetOptionTask extends Task, SheetCommonOptionTask {

		@Config("columns")
		@ConfigDefault("null")
		public Optional<Map<String, ColumnOptionTask>> getColumns();
	}

	public interface ColumnOptionTask extends Task, ColumnCommonOptionTask {

		/**
		 * @see PoiExcelColumnValueType
		 * @return value_type
		 */
		@Config("value")
		@ConfigDefault("null")
		public Optional<String> getValueType();

		// same as cell_column
		@Config("column_number")
		@ConfigDefault("null")
		public Optional<String> getColumnNumber();

		public static final String CELL_COLUMN = "cell_column";

		// A,B,... or number(1 origin)
		@Config(CELL_COLUMN)
		@ConfigDefault("null")
		public Optional<String> getCellColumn();

		public static final String CELL_ROW = "cell_row";

		// number(1 origin)
		@Config(CELL_ROW)
		@ConfigDefault("null")
		public Optional<String> getCellRow();

		// A1,B2,... or Sheet1!A1
		@Config("cell_address")
		@ConfigDefault("null")
		public Optional<String> getCellAddress();

		// use when value_type=cell_style, cell_font, ...
		@Config("attribute_name")
		@ConfigDefault("null")
		public Optional<List<String>> getAttributeName();
	}

	public interface ColumnCommonOptionTask extends Task {
		// format of numeric(double) to string
		@Config("numeric_format")
		@ConfigDefault("null")
		public Optional<String> getNumericFormat();

		// search merged cell if cellType=BLANK
		@Config("search_merged_cell")
		@ConfigDefault("null")
		public Optional<String> getSearchMergedCell();

		@Config("formula_handling")
		@ConfigDefault("null")
		public Optional<String> getFormulaHandling();

		@Config("formula_replace")
		@ConfigDefault("null")
		public Optional<List<FormulaReplaceTask>> getFormulaReplace();

		@Config("on_evaluate_error")
		@ConfigDefault("null")
		public Optional<String> getOnEvaluateError();

		@Config("on_cell_error")
		@ConfigDefault("null")
		public Optional<String> getOnCellError();

		@Config("on_convert_error")
		@ConfigDefault("null")
		public Optional<String> getOnConvertError();
	}

	public interface FormulaReplaceTask extends Task {

		@Config("regex")
		public String getRegex();

		// replace string
		// can use variable: "${row}"
		@Config("to")
		public String getTo();
	}

	@Override
	public void transaction(ConfigSource config, ParserPlugin.Control control) {
		PluginTask task = config.loadConfig(PluginTask.class);

		Schema schema = task.getColumns().toSchema();

		control.run(task.dump(), schema);
	}

	@Override
	public void run(TaskSource taskSource, Schema schema, FileInput input, PageOutput output) {
		PluginTask task = taskSource.loadTask(PluginTask.class);

		List<String> sheetNames = new ArrayList<>();
		Optional<String> sheetOption = task.getSheet();
		if (sheetOption.isPresent()) {
			sheetNames.add(sheetOption.get());
		}
		sheetNames.addAll(task.getSheets());
		if (sheetNames.isEmpty()) {
			throw new ConfigException("Attribute sheets is required but not set");
		}

		try (FileInputInputStream is = new FileInputInputStream(input)) {
			while (is.nextFile()) {
				Workbook workbook;
				try {
					workbook = WorkbookFactory.create(is);
				} catch (IOException | EncryptedDocumentException e) {
					throw new RuntimeException(e);
				}

				List<String> list = resolveSheetName(workbook, sheetNames);
				if (log.isDebugEnabled()) {
					log.debug("resolved sheet names={}", list);
				}
				run(task, schema, workbook, list, output);
			}
		}
	}

	private List<String> resolveSheetName(Workbook workbook, List<String> sheetNames) {
		Set<String> set = new LinkedHashSet<>();
		for (String s : sheetNames) {
			if (s.contains("*") || s.contains("?")) {
				int length = s.length();
				StringBuilder sb = new StringBuilder(length * 2);
				StringBuilder buf = new StringBuilder(32);
				for (int i = 0; i < length;) {
					int c = s.codePointAt(i);
					switch (c) {
					case '*':
						if (buf.length() > 0) {
							sb.append(Pattern.quote(buf.toString()));
							buf.setLength(0);
						}
						sb.append(".*");
						break;
					case '?':
						if (buf.length() > 0) {
							sb.append(Pattern.quote(buf.toString()));
							buf.setLength(0);
						}
						sb.append(".");
						break;
					default:
						buf.appendCodePoint(c);
						break;
					}
					i += Character.charCount(c);
				}
				if (buf.length() > 0) {
					sb.append(Pattern.quote(buf.toString()));
				}
				String regex = sb.toString();
				for (Sheet sheet : workbook) {
					String name = sheet.getSheetName();
					if (name.matches(regex)) {
						set.add(name);
					}
				}
			} else {
				set.add(s);
			}
		}
		return new ArrayList<>(set);
	}

	protected void run(PluginTask task, Schema schema, Workbook workbook, List<String> sheetNames, PageOutput output) {
		final int flushCount = task.getFlushCount();

		try (PageBuilder pageBuilder = new PageBuilder(Exec.getBufferAllocator(), schema, output)) {
			for (String sheetName : sheetNames) {
				Sheet sheet = workbook.getSheet(sheetName);
				if (sheet == null) {
					if (task.getIgnoreSheetNotFound()) {
						log.info("ignore: not found sheet={}", sheetName);
						continue;
					} else {
						throw new RuntimeException(String.format("not found sheet=%s", sheetName));
					}
				}

				log.info("sheet={}", sheetName);
				PoiExcelVisitorFactory factory = newPoiExcelVisitorFactory(task, schema, sheet, pageBuilder);
				PoiExcelColumnVisitor visitor = factory.getPoiExcelColumnVisitor();
				PoiExcelSheetBean sheetBean = factory.getVisitorValue().getSheetBean();
				final int skipHeaderLines = sheetBean.getSkipHeaderLines();

				PoiExcelRecord record = sheetBean.getRecordType().newPoiExcelRecord();
				record.initialize(sheet, skipHeaderLines);
				visitor.setRecord(record);

				int count = 0;
				for (; record.exists(); record.moveNext()) {
					record.logStart();

					schema.visitColumns(visitor); // use record
					pageBuilder.addRecord();

					if (++count >= flushCount) {
						log.trace("flush");
						pageBuilder.flush();
						count = 0;
					}

					record.logEnd();
				}
				pageBuilder.flush();
			}
			pageBuilder.finish();
		}
	}

	protected PoiExcelVisitorFactory newPoiExcelVisitorFactory(PluginTask task, Schema schema, Sheet sheet,
			PageBuilder pageBuilder) {
		PoiExcelVisitorValue visitorValue = new PoiExcelVisitorValue(task, schema, sheet, pageBuilder);
		return new PoiExcelVisitorFactory(visitorValue);
	}
}
