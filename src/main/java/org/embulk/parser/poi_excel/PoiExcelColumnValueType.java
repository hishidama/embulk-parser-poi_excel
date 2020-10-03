package org.embulk.parser.poi_excel;

import org.embulk.parser.poi_excel.bean.record.RecordType;

public enum PoiExcelColumnValueType {
	/** cell value */
	CELL_VALUE(true, true),
	/** cell formula */
	CELL_FORMULA(true, true),
	/** cell style */
	CELL_STYLE(true, false),
	/** cell font */
	CELL_FONT(true, false),
	/** cell comment */
	CELL_COMMENT(true, false),
	/** cell type */
	CELL_TYPE(true, false),
	/** cell CachedFormulaResultType */
	CELL_CACHED_TYPE(true, false),
	/** sheet name */
	SHEET_NAME(false, false),
	/** row number (1 origin) */
	ROW_NUMBER(false, false) {
		@Override
		public boolean useCell(RecordType recordType) {
			if (recordType == RecordType.COLUMN) {
				return true;
			}
			return super.useCell(recordType);
		}
	},
	/** column number (1 origin) */
	COLUMN_NUMBER(true, false) {
		@Override
		public boolean useCell(RecordType recordType) {
			if (recordType == RecordType.ROW) {
				return true;
			}
			return super.useCell(recordType);
		}
	},
	/** constant */
	CONSTANT(false, false);

	private final boolean useCell;
	private final boolean nextIndex;

	PoiExcelColumnValueType(boolean useCell, boolean nextIndex) {
		this.useCell = useCell;
		this.nextIndex = nextIndex;
	}

	public boolean useCell(RecordType recordType) {
		return useCell;
	}

	public boolean nextIndex() {
		return nextIndex;
	}
}
