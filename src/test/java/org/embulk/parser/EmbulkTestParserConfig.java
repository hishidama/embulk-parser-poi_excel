package org.embulk.parser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

@SuppressWarnings("serial")
public class EmbulkTestParserConfig extends HashMap<String, Object> {

	public void setType(String type) {
		set("type", type);
	}

	public void set(String key, Object value) {
		if (value == null) {
			super.remove(key);
		} else {
			super.put(key, value);
		}
	}

	public List<EmbulkTestColumn> getColumns() {
		@SuppressWarnings("unchecked")
		List<EmbulkTestColumn> columns = (List<EmbulkTestColumn>) super.get("columns");
		if (columns == null) {
			columns = new ArrayList<>();
			super.put("columns", columns);
		}
		return columns;
	}

	public EmbulkTestColumn addColumn(String name, String type) {
		EmbulkTestColumn column = new EmbulkTestColumn();
		column.set("name", name);
		column.set("type", type);
		getColumns().add(column);
		return column;
	}

	public static class EmbulkTestColumn extends HashMap<String, Object> {

		public EmbulkTestColumn set(String key, Object value) {
			if (value == null) {
				super.remove(key);
			} else {
				super.put(key, value);
			}
			return this;
		}
	}
}
