package org.embulk.parser.poi_excel;

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
	/** sheet name */
	SHEET_NAME(false, false),
	/** row number (1 origin) */
	ROW_NUMBER(false, false),
	/** column number (1 origin) */
	COLUMN_NUMBER(true, false),
	/** constant */
	CONSTANT(false, false);

	private final boolean useCell;
	private final boolean nextIndex;

	PoiExcelColumnValueType(boolean useCell, boolean nextIndex) {
		this.useCell = useCell;
		this.nextIndex = nextIndex;
	}

	public boolean useCell() {
		return useCell;
	}

	public boolean nextIndex() {
		return nextIndex;
	}
}
