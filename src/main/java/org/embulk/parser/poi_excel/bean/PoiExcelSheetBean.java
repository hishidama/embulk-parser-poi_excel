package org.embulk.parser.poi_excel.bean;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.SheetOptionTask;
import org.embulk.spi.Column;
import org.embulk.spi.ColumnConfig;
import org.embulk.spi.Schema;

import com.google.common.base.Optional;

public class PoiExcelSheetBean {

	protected final Sheet sheet;

	private final List<SheetOptionTask> sheetTaskList = new ArrayList<>(2);

	private final List<PoiExcelColumnBean> columnBeanList = new ArrayList<>();

	public PoiExcelSheetBean(PluginTask task, Schema schema, Sheet sheet) {
		this.sheet = sheet;

		initializeSheetTask(task);
		initializeColumnBean(task, schema);
	}

	private void initializeSheetTask(PluginTask task) {
		String name = sheet.getSheetName();
		Map<String, SheetOptionTask> map = task.getSheetOptions();
		SheetOptionTask s = map.get(name);
		if (s != null) {
			sheetTaskList.add(s);
		} else {
			loop: for (Entry<String, SheetOptionTask> entry : map.entrySet()) {
				String[] ss = entry.getKey().split("/");
				for (String key : ss) {
					if (key.trim().equalsIgnoreCase(name)) {
						sheetTaskList.add(entry.getValue());
						break loop;
					}
				}
			}
		}
		sheetTaskList.add(task);
	}

	private void initializeColumnBean(PluginTask task, Schema schema) {
		List<ColumnConfig> list = task.getColumns().getColumns();

		Map<String, ColumnOptionTask> map = new HashMap<>();
		List<SheetOptionTask> slist = getSheetOption();
		if (slist.size() >= 2) {
			SheetOptionTask s = slist.get(0);
			for (ColumnConfig c : s.getColumns().getColumns()) {
				String name = c.getName();
				ColumnOptionTask t = c.getOption().loadConfig(ColumnOptionTask.class);
				map.put(name, t);
			}
		}

		for (Column column : schema.getColumns()) {
			String name = column.getName();
			ColumnConfig c = list.get(column.getIndex());
			ColumnOptionTask t = c.getOption().loadConfig(ColumnOptionTask.class);
			PoiExcelColumnBean bean = new PoiExcelColumnBean(this, column, t, map.get(name));
			columnBeanList.add(bean);
		}

		new PoiExcelColumnIndex().initializeColumnIndex(task, columnBeanList);
	}

	public final List<SheetOptionTask> getSheetOption() {
		return sheetTaskList;
	}

	public int getSkipHeaderLines() {
		List<SheetOptionTask> list = getSheetOption();
		for (SheetOptionTask sheetTask : list) {
			Optional<Integer> value = sheetTask.getSkipHeaderLines();
			if (value.isPresent()) {
				return value.get();
			}
		}
		return 0;
	}

	public final List<PoiExcelColumnBean> getColumnBeans() {
		return columnBeanList;
	}

	public final PoiExcelColumnBean getColumnBean(Column column) {
		List<PoiExcelColumnBean> list = getColumnBeans();
		return list.get(column.getIndex());
	}
}
