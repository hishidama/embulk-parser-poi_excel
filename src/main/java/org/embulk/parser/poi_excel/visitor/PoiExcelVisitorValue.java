package org.embulk.parser.poi_excel.visitor;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.spi.Column;
import org.embulk.spi.ColumnConfig;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.SchemaConfig;

public class PoiExcelVisitorValue {
	private final PluginTask task;
	private final Sheet sheet;
	private final PageBuilder pageBuilder;
	private PoiExcelVisitorFactory factory;

	private List<ColumnOptionTask> columnOptions;

	public PoiExcelVisitorValue(PluginTask task, Sheet sheet, PageBuilder pageBuilder) {
		this.task = task;
		this.sheet = sheet;
		this.pageBuilder = pageBuilder;
	}

	public PluginTask getPluginTask() {
		return task;
	}

	public Sheet getSheet() {
		return sheet;
	}

	public PageBuilder getPageBuilder() {
		return pageBuilder;
	}

	public void setVisitorFactory(PoiExcelVisitorFactory factory) {
		this.factory = factory;
	}

	public PoiExcelVisitorFactory getVisitorFactory() {
		return factory;
	}

	public ColumnOptionTask getColumnOption(Column column) {
		return getColumnOptions().get(column.getIndex());
	}

	public List<ColumnOptionTask> getColumnOptions() {
		if (columnOptions == null) {
			SchemaConfig schemaConfig = task.getColumns();
			columnOptions = new ArrayList<>(schemaConfig.getColumnCount());
			for (ColumnConfig c : schemaConfig.getColumns()) {
				ColumnOptionTask option = c.getOption().loadConfig(ColumnOptionTask.class);
				columnOptions.add(option);
			}
		}
		return columnOptions;
	}
}
