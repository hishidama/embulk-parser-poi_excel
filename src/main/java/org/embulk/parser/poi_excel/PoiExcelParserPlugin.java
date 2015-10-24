package org.embulk.parser.poi_excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.embulk.config.Config;
import org.embulk.config.ConfigDefault;
import org.embulk.config.ConfigException;
import org.embulk.config.ConfigSource;
import org.embulk.config.Task;
import org.embulk.config.TaskSource;
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
import com.ibm.icu.text.MessageFormat;

public class PoiExcelParserPlugin implements ParserPlugin {
	private final Logger log = Exec.getLogger(getClass());

	public static final String TYPE = "poi_excel";

	public interface PluginTask extends Task, TimestampParser.Task, SheetOptionTask {
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

		@Config("flush_count")
		@ConfigDefault("100")
		public int getFlushCount();
	}

	public interface SheetOptionTask extends Task, ColumnCommonOptionTask {

		@Config("skip_header_lines")
		@ConfigDefault("null")
		public Optional<Integer> getSkipHeaderLines();

		@Config("columns")
		public SchemaConfig getColumns();
	}

	public interface ColumnOptionTask extends Task, ColumnCommonOptionTask {

		/**
		 * @see PoiExcelColumnValueType
		 * @return value_type
		 */
		@Config("value")
		@ConfigDefault("null")
		public Optional<String> getValueType();

		// A,B,... or number(1 origin)
		@Config("column_number")
		@ConfigDefault("null")
		public Optional<String> getColumnNumber();

		// use when value_type=cell_style, cell_font, ...
		@Config("attribute_name")
		@ConfigDefault("null")
		public Optional<List<String>> getAttributeName();
	}

	public interface ColumnCommonOptionTask extends Task {

		// search merged cell if cellType=BLANK
		@Config("search_merged_cell")
		@ConfigDefault("null")
		public Optional<Boolean> getSearchMergedCell();

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
				} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
					throw new RuntimeException(e);
				}

				run(task, schema, workbook, sheetNames, output);
			}
		}
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
						throw new RuntimeException(MessageFormat.format("not found sheet={0}", sheetName));
					}
				}

				log.info("sheet={}", sheetName);
				PoiExcelVisitorFactory factory = newPoiExcelVisitorFactory(task, schema, sheet, pageBuilder);
				PoiExcelColumnVisitor visitor = factory.getPoiExcelColumnVisitor();
				final int skipHeaderLines = factory.getVisitorValue().getSheetBean().getSkipHeaderLines();

				int count = 0;
				for (Row row : sheet) {
					if (row.getRowNum() < skipHeaderLines) {
						continue;
					}

					visitor.setRow(row);
					schema.visitColumns(visitor);
					pageBuilder.addRecord();

					if (++count >= flushCount) {
						pageBuilder.flush();
						count = 0;
					}
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
