package org.embulk.parser.poi_excel.bean;

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.embulk.parser.poi_excel.PoiExcelColumnValueType;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnCommonOptionTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.FormulaReplaceTask;
import org.embulk.spi.Column;

import com.google.common.base.Optional;

public class PoiExcelColumnBean {

	protected final PoiExcelSheetBean sheetBean;
	protected final Column column;
	protected final List<ColumnOptionTask> columnTaskList = new ArrayList<>();
	protected final List<ColumnCommonOptionTask> allTaskList = new ArrayList<>();

	private PoiExcelColumnValueType valueType;
	private String valueTypeSuffix;
	private int columnIndex;

	public PoiExcelColumnBean(PoiExcelSheetBean sheetBean, Column column, ColumnOptionTask mainTask,
			ColumnOptionTask optionTask) {
		this.sheetBean = sheetBean;
		this.column = column;
		if (optionTask != null) {
			columnTaskList.add(optionTask);
			allTaskList.add(optionTask);
		}
		columnTaskList.add(mainTask);
		allTaskList.add(mainTask);
		allTaskList.addAll(sheetBean.getSheetOption());

		initialize();
	}

	private void initialize() {
		String type = null;
		for (ColumnOptionTask task : columnTaskList) {
			Optional<String> option = task.getValueType();
			if (option.isPresent()) {
				type = option.get();
				break;
			}
		}
		if (type == null) {
			this.valueType = PoiExcelColumnValueType.CELL_VALUE;
			return;
		}

		int n = type.indexOf('.');
		if (n >= 0) {
			String suffix = type.substring(n + 1).trim();
			this.valueTypeSuffix = suffix;
			type = type.substring(0, n).trim();
		}

		try {
			this.valueType = PoiExcelColumnValueType.valueOf(type.toUpperCase());
		} catch (Exception e) {
			throw new RuntimeException(MessageFormat.format("illegal value_type={0}", type), e);
		}
	}

	public PoiExcelColumnValueType getValueType() {
		return valueType;
	}

	public String getValueTypeSuffix() {
		return valueTypeSuffix;
	}

	public void setColumnIndex(int columnIndex) {
		this.columnIndex = columnIndex;
	}

	public int getColumnIndex() {
		return columnIndex;
	}

	public Optional<String> getColumnNumber() {
		for (ColumnOptionTask task : columnTaskList) {
			Optional<String> option = task.getColumnNumber();
			if (option.isPresent()) {
				return option;
			}
		}
		return Optional.absent();
	}

	protected abstract class CacheValue<T> {
		private T value;

		public CacheValue() {
		}

		public T get() {
			if (value == null) {
				T v = null;
				for (ColumnCommonOptionTask task : allTaskList) {
					Optional<T> option = getTaskValue(task);
					if (option.isPresent()) {
						v = option.get();
						break;
					}
				}
				if (v == null) {
					v = getDefaultValue();
				}
				this.value = v;
			}
			return value;
		}

		protected abstract Optional<T> getTaskValue(ColumnCommonOptionTask task);

		protected abstract T getDefaultValue();
	}

	private CacheValue<List<String>> attributeName = new CacheValue<List<String>>() {

		@Override
		protected Optional<List<String>> getTaskValue(ColumnCommonOptionTask task) {
			if (task instanceof ColumnOptionTask) {
				return ((ColumnOptionTask) task).getAttributeName();
			}
			return Optional.absent();
		}

		@Override
		protected List<String> getDefaultValue() {
			return Collections.emptyList();
		}
	};

	public List<String> getAttributeName() {
		return attributeName.get();
	}

	private CacheValue<Boolean> searchMergedCell = new CacheValue<Boolean>() {

		@Override
		protected Optional<Boolean> getTaskValue(ColumnCommonOptionTask task) {
			return task.getSearchMergedCell();
		}

		@Override
		protected Boolean getDefaultValue() {
			return true;
		}
	};

	public boolean getSearchMergedCell() {
		return searchMergedCell.get();
	}

	private CacheValue<List<FormulaReplaceTask>> formulaReplace = new CacheValue<List<FormulaReplaceTask>>() {

		@Override
		protected Optional<List<FormulaReplaceTask>> getTaskValue(ColumnCommonOptionTask task) {
			return task.getFormulaReplace();
		}

		@Override
		protected List<FormulaReplaceTask> getDefaultValue() {
			return Collections.emptyList();
		}
	};

	public List<FormulaReplaceTask> getFormulaReplace() {
		return formulaReplace.get();
	}

	private CacheValue<Boolean> formulaErrorNull = new CacheValue<Boolean>() {

		@Override
		protected Optional<Boolean> getTaskValue(ColumnCommonOptionTask task) {
			return task.getFormulaErrorNull();
		}

		@Override
		protected Boolean getDefaultValue() {
			return false;
		}
	};

	public boolean getFormulaErrorNull() {
		return formulaErrorNull.get();
	}

	private CacheValue<Boolean> cellErrorNull = new CacheValue<Boolean>() {

		@Override
		protected Optional<Boolean> getTaskValue(ColumnCommonOptionTask task) {
			return task.getCellErrorNull();
		}

		@Override
		protected Boolean getDefaultValue() {
			return true;
		}
	};

	public boolean getCellErrorNull() {
		return cellErrorNull.get();
	}
}
