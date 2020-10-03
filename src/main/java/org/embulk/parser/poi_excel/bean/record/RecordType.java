package org.embulk.parser.poi_excel.bean.record;

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.List;

import org.embulk.config.ConfigException;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;

import com.google.common.base.Optional;

public enum RecordType {
	ROW {
		@Override
		public Optional<String> getRecordOption(PoiExcelColumnBean bean) {
			return bean.getRowNumber();
		}

		@Override
		public String getRecordOptionName() {
			return ColumnOptionTask.CELL_ROW;
		}

		@Override
		public Optional<String> getNumberOption(PoiExcelColumnBean bean) {
			return bean.getColumnNumber();
		}

		@Override
		public String getNumberOptionName() {
			return ColumnOptionTask.CELL_COLUMN;
		}

		@Override
		public PoiExcelRecord newPoiExcelRecord() {
			return new PoiExcelRecordRow();
		}
	},
	COLUMN {
		@Override
		public Optional<String> getRecordOption(PoiExcelColumnBean bean) {
			return bean.getColumnNumber();
		}

		@Override
		public String getRecordOptionName() {
			return ColumnOptionTask.CELL_COLUMN;
		}

		@Override
		public Optional<String> getNumberOption(PoiExcelColumnBean bean) {
			return bean.getRowNumber();
		}

		@Override
		public String getNumberOptionName() {
			return ColumnOptionTask.CELL_ROW;
		}

		@Override
		public PoiExcelRecord newPoiExcelRecord() {
			return new PoiExcelRecordColumn();
		}
	},
	SHEET {
		@Override
		public Optional<String> getRecordOption(PoiExcelColumnBean bean) {
			return Optional.absent();
		}

		@Override
		public String getRecordOptionName() {
			return "-";
		}

		@Override
		public Optional<String> getNumberOption(PoiExcelColumnBean bean) {
			return Optional.absent();
		}

		@Override
		public String getNumberOptionName() {
			return "-";
		}

		@Override
		public PoiExcelRecord newPoiExcelRecord() {
			return new PoiExcelRecordSheet();
		}
	};

	public abstract Optional<String> getRecordOption(PoiExcelColumnBean bean);

	public abstract String getRecordOptionName();

	public abstract Optional<String> getNumberOption(PoiExcelColumnBean bean);

	public abstract String getNumberOptionName();

	public abstract PoiExcelRecord newPoiExcelRecord();

	public static RecordType of(String value) {
		try {
			return RecordType.valueOf(value.toUpperCase());
		} catch (Exception e) {
			List<String> list = new ArrayList<>();
			for (RecordType s : RecordType.values()) {
				list.add(s.name().toLowerCase());
			}
			throw new ConfigException(MessageFormat.format("illegal record_type={0}. expected={1}", value, list), e);
		}
	}
}
