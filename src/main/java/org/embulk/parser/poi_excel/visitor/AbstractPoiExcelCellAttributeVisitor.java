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
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.type.StringType;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.common.base.Optional;

public abstract class AbstractPoiExcelCellAttributeVisitor<A> {

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public AbstractPoiExcelCellAttributeVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visit(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
		A source = getAttributeSource(column, option, cell);
		if (source == null) {
			pageBuilder.setNull(column);
			return;
		}

		String suffix = option.getValueTypeSuffix();
		if (suffix != null) {
			visitKey(column, option, suffix, cell, source, visitor);
		} else {
			visitJson(column, option, cell, source, visitor);
		}
	}

	protected abstract A getAttributeSource(Column column, ColumnOptionTask option, Cell cell);

	private void visitKey(Column column, ColumnOptionTask option, String key, Cell cell, A source, CellVisitor visitor) {
		Object value = getAttributeValue(column, option, cell, source, key);
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
		} else {
			throw new IllegalStateException(MessageFormat.format("unsupported conversion. type={0}, value={1}", value
					.getClass().getName(), value));
		}
	}

	private void visitJson(Column column, ColumnOptionTask option, Cell cell, A source, CellVisitor visitor) {
		Map<String, Object> result;

		Optional<List<String>> nameOption = option.getAttributeName();
		if (nameOption.isPresent()) {
			result = new LinkedHashMap<>();

			List<String> list = nameOption.get();
			for (String key : list) {
				Object value = getAttributeValue(column, option, cell, source, key);
				result.put(key, value);
			}
		} else {
			result = new TreeMap<>();

			Collection<String> keys = getAttributeSupplierMap().keySet();
			for (String key : keys) {
				if (acceptKey(key)) {
					Object value = getAttributeValue(column, option, cell, source, key);
					result.put(key, value);
				}
			}
		}

		String json;
		try {
			ObjectMapper mapper = new ObjectMapper();
			json = mapper.writeValueAsString(result);
		} catch (JsonProcessingException e) {
			throw new RuntimeException(e);
		}

		visitor.visitCellValueString(column, cell, json);
	}

	protected boolean acceptKey(String key) {
		return true;
	}

	private Object getAttributeValue(Column column, ColumnOptionTask option, Cell cell, A source, String key) {
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
}
