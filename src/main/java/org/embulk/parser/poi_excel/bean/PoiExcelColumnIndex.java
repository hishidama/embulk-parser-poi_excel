package org.embulk.parser.poi_excel.bean;

import java.text.MessageFormat;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.util.CellReference;
import org.embulk.parser.poi_excel.PoiExcelColumnValueType;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.parser.poi_excel.bean.record.RecordType;
import org.embulk.parser.poi_excel.bean.util.PoiExcelCellAddress;
import org.embulk.spi.Column;
import org.embulk.spi.Exec;
import org.embulk.spi.Schema;
import org.slf4j.Logger;

import com.google.common.base.Optional;

public class PoiExcelColumnIndex {
	private final Logger log = Exec.getLogger(getClass());

	protected final RecordType recordType;
	protected final Map<String, Integer> indexMap = new LinkedHashMap<>();

	public PoiExcelColumnIndex(PoiExcelSheetBean sheetBean) {
		this.recordType = sheetBean.getRecordType();
	}

	public void initializeColumnIndex(PluginTask task, List<PoiExcelColumnBean> beanList) {
		log.info("record_type={}", recordType);

		int index = -1;
		indexMap.clear();

		Schema schema = task.getColumns().toSchema();
		for (Column column : schema.getColumns()) {
			PoiExcelColumnBean bean = beanList.get(column.getIndex());
			initializeCellAddress(column, bean);

			PoiExcelColumnValueType valueType = bean.getValueType();
			if (valueType.useCell(recordType)) {
				index = resolveColumnIndex(column, bean, index, valueType);
				if (index < 0) {
					index = 0;
				}
				bean.setColumnIndex(index);
				indexMap.put(column.getName(), index);
			}

			initializeCellAddress2(column, bean, index);

			if (log.isInfoEnabled()) {
				logColumn(column, bean, valueType, index);
			}
		}
	}

	protected void initializeCellAddress(Column column, PoiExcelColumnBean bean) {
		if (bean.getCellAddress() != null) {
			return;
		}

		Optional<String> rowOption = bean.getRowNumber();
		Optional<String> colOption = bean.getColumnNumber();
		if (rowOption.isPresent() && colOption.isPresent()) {
			String rowNumber = rowOption.get();
			String colNumber = colOption.get();
			initializeCellAddress(column, bean, rowNumber, colNumber);
			return;
		}

		if (recordType == RecordType.SHEET) {
			String rowNumber = rowOption.or("1");
			String colNumber = colOption.or("A");
			initializeCellAddress(column, bean, rowNumber, colNumber);
			return;
		}
	}

	protected void initializeCellAddress(Column column, PoiExcelColumnBean bean, String rowNumber, String columnNumber) {
		int columnIndex = convertColumnIndex(column, OPTION_NAME_CELL_COLUMN, columnNumber);
		int rowIndex = convertColumnIndex(column, OPTION_NAME_CELL_ROW, rowNumber);
		CellReference ref = new CellReference(rowIndex, columnIndex);
		bean.setCellAddress(ref);
	}

	protected void initializeCellAddress2(Column column, PoiExcelColumnBean bean, int index) {
		if (bean.getCellAddress() != null) {
			return;
		}

		Optional<String> recordOption = recordType.getRecordOption(bean);
		if (recordOption.isPresent()) {
			int rowIndex, columnIndex;
			switch (recordType) {
			case ROW:
				rowIndex = convertColumnIndex(column, recordType.getRecordOptionName(), recordOption.get());
				columnIndex = (index >= 0) ? index : 0;
				break;
			case COLUMN:
				rowIndex = (index >= 0) ? index : 0;
				columnIndex = convertColumnIndex(column, recordType.getRecordOptionName(), recordOption.get());
				break;
			default:
				throw new IllegalStateException();
			}
			CellReference ref = new CellReference(rowIndex, columnIndex);
			bean.setCellAddress(ref);
		}
	}

	protected int resolveColumnIndex(Column column, PoiExcelColumnBean bean, int index,
			PoiExcelColumnValueType valueType) {
		Optional<String> numberOption = recordType.getNumberOption(bean);
		PoiExcelCellAddress cellAddress = bean.getCellAddress();

		if (cellAddress != null) {
			return index;
		}

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
			return convertColumnIndex(column, recordType.getNumberOptionName(), columnNumber);
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
			throw new RuntimeException(MessageFormat.format("{0} out of range at {1}",
					recordType.getNumberOptionName(), column));
		}
	}

	protected int convertColumnIndex(Column column, String numberOptionName, String columnNumber) {
		int index;
		try {
			char c = columnNumber.charAt(0);
			if ('0' <= c && c <= '9') {
				index = Integer.parseInt(columnNumber) - 1;
			} else {
				index = CellReference.convertColStringToIndex(columnNumber);
			}
		} catch (Exception e) {
			throw new RuntimeException(MessageFormat.format("illegal {0}=\"{1}\" at {2}", numberOptionName,
					columnNumber, column), e);
		}
		if (index < 0) {
			throw new RuntimeException(MessageFormat.format("illegal {0}=\"{1}\" at {2}", numberOptionName,
					columnNumber, column));
		}
		return index;
	}

	private static final String OPTION_NAME_CELL_COLUMN = ColumnOptionTask.CELL_COLUMN;
	private static final String OPTION_NAME_CELL_ROW = ColumnOptionTask.CELL_ROW;

	protected void logColumn(Column column, PoiExcelColumnBean bean, PoiExcelColumnValueType valueType, int index) {
		PoiExcelCellAddress cellAddress = bean.getCellAddress();

		String cname, cvalue;
		if (cellAddress != null) {
			cname = "cell_address";
			cvalue = cellAddress.getString();
		} else {
			switch (recordType) {
			default:
				cname = OPTION_NAME_CELL_COLUMN;
				cvalue = CellReference.convertNumToColString(index);
				break;
			case COLUMN:
				cname = OPTION_NAME_CELL_ROW;
				cvalue = Integer.toString(index + 1);
				break;
			case SHEET:
				cname = "sheet";
				cvalue = null;
				break;
			}
		}

		switch (valueType) {
		default:
		case CELL_VALUE:
		case CELL_FORMULA:
		case CELL_TYPE:
		case CELL_CACHED_TYPE:
			log.info("column.name={} <- {}={}, value={}", column.getName(), cname, cvalue, valueType);
			break;
		case CELL_STYLE:
		case CELL_FONT:
		case CELL_COMMENT:
			String suffix = bean.getValueTypeSuffix();
			if (suffix != null) {
				log.info("column.name={} <- {}={}, value={}[{}]", column.getName(), cname, cvalue, valueType, suffix);
			} else {
				log.info("column.name={} <- {}={}, value={}", column.getName(), cname, cvalue, valueType);
			}
			break;

		case SHEET_NAME:
			if (cellAddress != null && cellAddress.getSheetName() != null) {
				log.info("column.name={} <- {}={}, value={}", column.getName(), cname, cvalue, valueType);
			} else {
				log.info("column.name={} <- value={}", column.getName(), valueType);
			}
			break;
		case ROW_NUMBER:
			if (cellAddress != null || cname.equals(OPTION_NAME_CELL_ROW)) {
				log.info("column.name={} <- {}={}, value={}", column.getName(), cname, cvalue, valueType);
			} else {
				log.info("column.name={} <- value={}", column.getName(), valueType);
			}
			break;
		case COLUMN_NUMBER:
			if (cellAddress != null || cname.equals(OPTION_NAME_CELL_COLUMN)) {
				log.info("column.name={} <- {}={}, value={}", column.getName(), cname, cvalue, valueType);
			} else {
				log.info("column.name={} <- value={}", column.getName(), valueType);
			}
			break;

		case CONSTANT:
			String value = bean.getValueTypeSuffix();
			if (value != null) {
				log.info("column.name={} <- value={}[{}]", column.getName(), valueType, value);
			} else {
				log.info("column.name={} <- value={}({})", column.getName(), valueType, value);
			}
			break;
		}
	}
}
