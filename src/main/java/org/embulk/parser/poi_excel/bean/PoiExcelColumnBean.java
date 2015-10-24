package org.embulk.parser.poi_excel.bean;

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.embulk.config.ConfigException;
import org.embulk.parser.poi_excel.PoiExcelColumnValueType;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnCommonOptionTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.FormulaReplaceTask;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean.ErrorStrategy.Strategy;
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
		}
		columnTaskList.add(mainTask);

		allTaskList.addAll(columnTaskList);
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

		String suffix = null;
		{
			int n = type.indexOf('.');
			if (n >= 0) {
				suffix = type.substring(n + 1); // not trim
				this.valueTypeSuffix = suffix.trim();
				type = type.substring(0, n).trim();
			}
		}

		try {
			this.valueType = PoiExcelColumnValueType.valueOf(type.toUpperCase());
		} catch (Exception e) {
			throw new ConfigException(MessageFormat.format("illegal value_type={0}", type), e);
		}

		if (valueType == PoiExcelColumnValueType.CONSTANT) {
			this.valueTypeSuffix = suffix; // not trim
		}
	}

	public Column getColumn() {
		return column;
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

	public static final class ErrorStrategy {
		private final Strategy strategy;
		private final String value;

		public static enum Strategy {
			DEFAULT, EXCEPTION, CONSTANT, ERROR_CODE
		}

		public ErrorStrategy(Strategy strategy) {
			this.strategy = strategy;
			this.value = null;
		}

		public ErrorStrategy(String value) {
			this.strategy = Strategy.CONSTANT;
			this.value = value;
		}

		public Strategy getStrategy() {
			return strategy;
		}

		public String getValue() {
			return value;
		}

		@Override
		public String toString() {
			return String.format("ErrorStrategy(%s, %s)", strategy, value);
		}
	}

	protected abstract class CacheErrorStrategy extends CacheValue<ErrorStrategy> {

		public CacheErrorStrategy() {
		}

		@Override
		protected Optional<ErrorStrategy> getTaskValue(ColumnCommonOptionTask task) {
			Optional<String> option = getStringValue(task);
			if (!option.isPresent()) {
				return Optional.absent();
			}
			String value = option.get();
			if ("null".equalsIgnoreCase(value)) {
				value = Strategy.CONSTANT.name();
			}

			String suffix = null;
			int n = value.indexOf('.');
			if (n >= 0) {
				suffix = value.substring(n + 1);
				value = value.substring(0, n).trim();
			}
			try {
				Strategy strategy = Strategy.valueOf(value.toUpperCase());
				switch (strategy) {
				case CONSTANT:
					return Optional.of(new ErrorStrategy(suffix));
				default:
					return Optional.of(new ErrorStrategy(strategy));
				}
			} catch (Exception e) {
				throw new ConfigException(MessageFormat.format("illegal on-error type={0}", value), e);
			}
		}

		protected abstract Optional<String> getStringValue(ColumnCommonOptionTask task);

		@Override
		protected ErrorStrategy getDefaultValue() {
			return new ErrorStrategy(Strategy.DEFAULT);
		}
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

	private CacheErrorStrategy evaluateErrorStrategy = new CacheErrorStrategy() {
		@Override
		protected Optional<String> getStringValue(ColumnCommonOptionTask task) {
			return task.getOnEvaluateError();
		}
	};

	public ErrorStrategy getEvaluateErrorStrategy() {
		return evaluateErrorStrategy.get();
	}

	private CacheErrorStrategy cellErrorStrategy = new CacheErrorStrategy() {
		@Override
		protected Optional<String> getStringValue(ColumnCommonOptionTask task) {
			return task.getOnCellError();
		}
	};

	public ErrorStrategy getCellErrorStrategy() {
		return cellErrorStrategy.get();
	}

	private CacheErrorStrategy convertErrorStrategy = new CacheErrorStrategy() {
		@Override
		protected Optional<String> getStringValue(ColumnCommonOptionTask task) {
			return task.getOnConvertError();
		}
	};

	public ErrorStrategy getConvertErrorStrategy() {
		return convertErrorStrategy.get();
	}
}
