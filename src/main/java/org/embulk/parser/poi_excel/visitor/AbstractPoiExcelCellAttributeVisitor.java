package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Color;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.type.StringType;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

public abstract class AbstractPoiExcelCellAttributeVisitor<A> {

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public AbstractPoiExcelCellAttributeVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visit(Column column, PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
		A source = getAttributeSource(column, bean, cell);
		if (source == null) {
			pageBuilder.setNull(column);
			return;
		}

		String suffix = bean.getValueTypeSuffix();
		if (suffix != null) {
			visitKey(column, bean, suffix, cell, source, visitor);
		} else {
			visitJson(column, bean, cell, source, visitor);
		}
	}

	protected abstract A getAttributeSource(Column column, PoiExcelColumnBean bean, Cell cell);

	private void visitKey(Column column, PoiExcelColumnBean bean, String key, Cell cell, A source, CellVisitor visitor) {
		Object value = getAttributeValue(column, cell, source, key);
		if (value == null) {
			pageBuilder.setNull(column);
		} else if (value instanceof String) {
			visitor.visitCellValueString(column, source, (String) value);
		} else if (value instanceof Long) {
			visitor.visitValueLong(column, source, (Long) value);
		} else if (value instanceof Boolean) {
			visitor.visitCellValueBoolean(column, source, (Boolean) value);
		} else if (value instanceof Double) {
			visitor.visitCellValueNumeric(column, source, (Double) value);
		} else if (value instanceof Map) {
			visitor.visitCellValueString(column, source, convertJsonString(value));
		} else {
			throw new IllegalStateException(MessageFormat.format("unsupported conversion. type={0}, value={1}", value
					.getClass().getName(), value));
		}
	}

	private void visitJson(Column column, PoiExcelColumnBean bean, Cell cell, A source, CellVisitor visitor) {
		Map<String, Object> result;

		List<String> list = bean.getAttributeName();
		if (!list.isEmpty()) {
			result = getSpecifiedValues(column, cell, source, list);
		} else {
			result = getAllValues(column, cell, source);
		}

		String json = convertJsonString(result);
		visitor.visitCellValueString(column, cell, json);
	}

	protected final Map<String, Object> getSpecifiedValues(Column column, Cell cell, A source, List<String> keyList) {
		Map<String, Object> result = new LinkedHashMap<>();

		for (String key : keyList) {
			Object value = getAttributeValue(column, cell, source, key);
			result.put(key, value);
		}

		return result;
	}

	protected final Map<String, Object> getAllValues(Column column, Cell cell, A source) {
		Map<String, Object> result = new TreeMap<>();

		Collection<String> keys = getAttributeSupplierMap().keySet();
		for (String key : keys) {
			if (acceptKey(key)) {
				Object value = getAttributeValue(column, cell, source, key);
				result.put(key, value);
			}
		}

		return result;
	}

	protected boolean acceptKey(String key) {
		return true;
	}

	protected final Object getAttributeValue(Column column, Cell cell, A source, String key) {
		Map<String, AttributeSupplier<A>> map = getAttributeSupplierMap();
		AttributeSupplier<A> supplier = map.get(key.toLowerCase());
		if (supplier == null) {
			throw new UnsupportedOperationException(MessageFormat.format(
					"unsupported attribute name={0}, choose in {1}", key, new TreeSet<>(map.keySet())));
		}
		Object value = supplier.get(column, cell, source);

		if (value instanceof Color) {
			int rgb = PoiExcelColorVisitor.getRGB((Color) value);
			if (column.getType() instanceof StringType) {
				value = String.format("%06x", rgb);
			} else {
				value = (long) rgb;
			}
		}
		return value;
	}

	// @FunctionalInterface
	protected static interface AttributeSupplier<A> {
		public Object get(Column column, Cell cell, A source);
	}

	protected abstract Map<String, AttributeSupplier<A>> getAttributeSupplierMap();

	protected final String convertJsonString(Object result) {
		try {
			ObjectMapper mapper = new ObjectMapper();
			return mapper.writeValueAsString(result);
		} catch (JsonProcessingException e) {
			throw new RuntimeException(e);
		}
	}
}
