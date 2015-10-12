package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;

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

			Collection<String> keys = geyAllKeys();
			for (String key : keys) {
				Object value = getAttributeValue(column, option, cell, source, key);
				result.put(key, value);
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

	protected abstract Collection<String> geyAllKeys();

	protected abstract Object getAttributeValue(Column column, ColumnOptionTask option, Cell cell, A source, String key);
}
