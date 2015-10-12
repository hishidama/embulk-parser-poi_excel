package org.embulk.parser.poi_excel;

import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.embulk.config.Config;
import org.embulk.config.ConfigDefault;
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

import com.google.common.base.Optional;
import com.ibm.icu.text.MessageFormat;

public class PoiExcelParserPlugin implements ParserPlugin {

	public static final String TYPE = "poi_excel";

	public interface PluginTask extends Task, TimestampParser.Task {
		@Config("sheet")
		@ConfigDefault("\"Sheet1\"")
		public String getSheet();

		@Config("skip_header_lines")
		@ConfigDefault("0")
		public int getSkipHeaderLines();

		// search merged cell if cellType=BLANK
		@Config("search_merged_cell")
		@ConfigDefault("true")
		public boolean getSearchMergedCell();

		@Config("formula_replace")
		@ConfigDefault("null")
		public Optional<List<FormulaReplaceTask>> getFormulaReplace();

		// true: set null if formula error
		// false: throw exception if formula error
		@Config("formula_error_null")
		@ConfigDefault("false")
		public boolean getFormulaErrorNull();

		// true: set null if cell value error
		// false: set error code if cell value error
		@Config("cell_error_null")
		@ConfigDefault("true")
		public boolean getCellErrorNull();

		@Config("flush_count")
		@ConfigDefault("100")
		public int getFlushCount();

		@Config("columns")
		public SchemaConfig getColumns();
	}

	public interface ColumnOptionTask extends Task {

		/**
		 * @see PoiExcelColumnValueType
		 * @return value_type
		 */
		@Config("value")
		@ConfigDefault("\"cell_value\"")
		public String getValueType();

		public void setValueTypeEnum(PoiExcelColumnValueType valueType);

		public PoiExcelColumnValueType getValueTypeEnum();

		public void setValueTypeSuffix(String suffix);

		public String getValueTypeSuffix();

		// A,B,... or number(1 origin)
		@Config("column_number")
		@ConfigDefault("null")
		public Optional<String> getColumnNumber();

		public void setColumnIndex(int index);

		public int getColumnIndex();

		@Config("search_merged_cell")
		@ConfigDefault("null")
		public Optional<Boolean> getSearchMergedCell();

		@Config("formula_replace")
		@ConfigDefault("null")
		public Optional<List<FormulaReplaceTask>> getFormulaReplace();

		@Config("formula_error_null")
		@ConfigDefault("null")
		public Optional<Boolean> getFormulaErrorNull();

		@Config("cell_error_null")
		@ConfigDefault("null")
		public Optional<Boolean> getCellErrorNull();

		// use when value_type=cell_style, cell_font, ...
		@Config("attribute_name")
		@ConfigDefault("null")
		public Optional<List<String>> getAttributeName();
	}

	public interface FormulaReplaceTask extends Task {

		@Config("regex")
		public String getRegex();

		// replace string
		// use variable: "${row}"
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

		try (FileInputInputStream is = new FileInputInputStream(input)) {
			while (is.nextFile()) {
				Workbook workbook;
				try {
					workbook = WorkbookFactory.create(is);
				} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
					throw new RuntimeException(e);
				}

				String sheetName = task.getSheet();
				Sheet sheet = workbook.getSheet(sheetName);
				if (sheet == null) {
					throw new RuntimeException(MessageFormat.format("not found sheet={0}", sheetName));
				}

				run(task, schema, sheet, output);
			}
		}
	}

	protected void run(PluginTask task, Schema schema, Sheet sheet, PageOutput output) {
		int skipHeaderLines = task.getSkipHeaderLines();
		final int flushCount = task.getFlushCount();

		try (final PageBuilder pageBuilder = new PageBuilder(Exec.getBufferAllocator(), schema, output)) {
			PoiExcelVisitorFactory factory = newPoiExcelVisitorFactory(task, sheet, pageBuilder);
			PoiExcelColumnVisitor visitor = factory.getPoiExcelColumnVisitor();

			int count = 0;
			for (final Row row : sheet) {
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
			pageBuilder.finish();
		}
	}

	protected PoiExcelVisitorFactory newPoiExcelVisitorFactory(PluginTask task, Sheet sheet, PageBuilder pageBuilder) {
		PoiExcelVisitorValue visitorValue = new PoiExcelVisitorValue(task, sheet, pageBuilder);
		return new PoiExcelVisitorFactory(visitorValue);
	}
}
