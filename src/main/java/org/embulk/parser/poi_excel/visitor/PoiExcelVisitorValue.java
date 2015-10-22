package org.embulk.parser.poi_excel.visitor;

import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.parser.poi_excel.bean.PoiExcelSheetBean;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.Schema;

public class PoiExcelVisitorValue {
	private final PluginTask task;
	private final Sheet sheet;
	private final PageBuilder pageBuilder;
	private final PoiExcelSheetBean sheetBean;
	private PoiExcelVisitorFactory factory;

	public PoiExcelVisitorValue(PluginTask task, Schema schema, Sheet sheet, PageBuilder pageBuilder) {
		this.task = task;
		this.sheet = sheet;
		this.pageBuilder = pageBuilder;
		this.sheetBean = new PoiExcelSheetBean(task, schema, sheet);
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

	public PoiExcelSheetBean getSheetBean() {
		return sheetBean;
	}

	public PoiExcelColumnBean getColumnBean(Column column) {
		return sheetBean.getColumnBean(column);
	}
}
