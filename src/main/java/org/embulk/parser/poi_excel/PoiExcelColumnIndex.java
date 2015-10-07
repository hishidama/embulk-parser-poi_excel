package org.embulk.parser.poi_excel;

import java.text.MessageFormat;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.util.CellReference;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.spi.Column;
import org.embulk.spi.Exec;
import org.embulk.spi.Schema;
import org.slf4j.Logger;

import com.google.common.base.Optional;

public class PoiExcelColumnIndex {
	private final Logger log = Exec.getLogger(getClass());

	protected final Map<String, Integer> indexMap = new LinkedHashMap<>();

	public void initializeColumnIndex(PluginTask task, List<ColumnOptionTask> columnOptions) {
		int index = -1;
		indexMap.clear();

		Schema schema = task.getColumns().toSchema();
		for (Column column : schema.getColumns()) {
			ColumnOptionTask option = columnOptions.get(column.getIndex());

			PoiExcelColumnValueType valueType;
			if (option.getCellStyleName().isPresent()) {
				valueType = PoiExcelColumnValueType.CELL_STYLE;
			} else {
				String type = option.getValueType();
				try {
					valueType = PoiExcelColumnValueType.valueOf(type.toUpperCase());
				} catch (Exception e) {
					throw new RuntimeException(MessageFormat.format("illegal value_type={0}", type), e);
				}
			}
			option.setValueTypeEnum(valueType);

			if (valueType.useCell()) {
				index = resolveColumnIndex(column, option, index, valueType);
				if (index < 0) {
					index = 0;
				}
				log.info("column.name={} <- cell column={}", column.getName(),
						CellReference.convertNumToColString(index));
				option.setColumnIndex(index);
				indexMap.put(column.getName(), index);
			}
		}
	}

	protected int resolveColumnIndex(Column column, ColumnOptionTask option, int index,
			PoiExcelColumnValueType valueType) {
		Optional<String> numberOption = option.getColumnNumber();
		if (numberOption.isPresent()) {
			String columnNumber = numberOption.get();
			if (columnNumber.length() >= 1) {
				char c = columnNumber.charAt(0);
				String arg = columnNumber.substring(1).trim();
				switch (c) {
				case '=':
					return resolveSameColumnIndex(column, index, columnNumber, arg);
				case '+':
					return resolveNextColumnIndex(column, index, columnNumber, arg);
				case '-':
					return resolvePreviousColumnIndex(column, index, columnNumber, arg);
				default:
					break;
				}
			}
			return convertColumnIndex(column, columnNumber);
		} else {
			if (valueType.nextIndex()) {
				index++;
			}
			return index;
		}
	}

	protected int resolveSameColumnIndex(Column column, int index, String columnNumber, String arg) {
		if (arg.isEmpty()) {
			return index;
		}

		Integer value = indexMap.get(arg);
		if (value == null) {
			throw new RuntimeException(MessageFormat.format("not found column name={0} before {1}", arg, column));
		}
		return value;
	}

	protected int resolveNextColumnIndex(Column column, int index, String columnNumber, String arg) {
		if (index < 0) {
			index = 0;
		}
		int add = 1;
		if (!arg.isEmpty()) {
			try {
				add = Integer.parseInt(arg);
			} catch (Exception e) {
				Integer value = indexMap.get(arg);
				if (value == null) {
					throw new RuntimeException(
							MessageFormat.format("not found column name={0} before {1}", arg, column));
				}
				index = value;
				add = 1;
			}
		}

		index += add;
		checkIndex(column, index);
		return index;
	}

	protected int resolvePreviousColumnIndex(Column column, int index, String columnNumber, String arg) {
		if (index < 0) {
			index = 0;
		}
		int sub = 1;
		if (!arg.isEmpty()) {
			try {
				sub = Integer.parseInt(arg);
			} catch (Exception e) {
				Integer value = indexMap.get(arg);
				if (value == null) {
					throw new RuntimeException(
							MessageFormat.format("not found column name={0} before {1}", arg, column));
				}
				index = value;
				sub = 1;
			}
		}

		index -= sub;
		checkIndex(column, index);
		return index;
	}

	protected void checkIndex(Column column, int index) {
		if (index < 0) {
			throw new RuntimeException(MessageFormat.format("column_number out of range. {0}", column));
		}
	}

	protected int convertColumnIndex(Column column, String columnNumber) {
		int index;
		try {
			char c = columnNumber.charAt(0);
			if ('0' <= c && c <= '9') {
				index = Integer.parseInt(columnNumber) - 1;
			} else {
				index = CellReference.convertColStringToIndex(columnNumber);
			}
		} catch (Exception e) {
			throw new RuntimeException(
					MessageFormat.format("illegal column_number=\"{0}\". {1}", columnNumber, column), e);
		}
		if (index < 0) {
			throw new RuntimeException(MessageFormat.format("illegal column_number=\"{0}\". {1}", columnNumber, column));
		}
		return index;
	}
}
