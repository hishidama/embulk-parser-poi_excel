package org.embulk.parser.poi_excel.bean.record;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;

public abstract class PoiExcelRecord {

	// loop record

	private Sheet sheet;

	public final void initialize(Sheet sheet, int skipHeaderLines) {
		this.sheet = sheet;
		initializeLoop(skipHeaderLines);
	}

	protected abstract void initializeLoop(int skipHeaderLines);

	public final Sheet getSheet() {
		return sheet;
	}

	public abstract boolean exists();

	public abstract void moveNext();

	// current record

	public final void logStart() {
		logStartEnd("start");
	}

	public final void logEnd() {
		logStartEnd("end");
	}

	protected abstract void logStartEnd(String part);

	public abstract int getRowIndex(PoiExcelColumnBean bean);

	public abstract int getColumnIndex(PoiExcelColumnBean bean);

	public abstract Cell getCell(PoiExcelColumnBean bean);

	public CellReference getCellReference(PoiExcelColumnBean bean) {
		int rowIndex = getRowIndex(bean);
		int columnIndex = getColumnIndex(bean);
		return new CellReference(rowIndex, columnIndex);
	}
}
