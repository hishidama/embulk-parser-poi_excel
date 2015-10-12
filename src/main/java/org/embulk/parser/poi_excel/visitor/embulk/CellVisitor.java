package org.embulk.parser.poi_excel.visitor.embulk;

import org.apache.poi.ss.usermodel.Cell;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorValue;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;

public abstract class CellVisitor {

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public CellVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public abstract void visitCellValueNumeric(Column column, Object source, double value);

	public abstract void visitCellValueString(Column column, Object source, String value);

	public void visitCellValueBlank(Column column, Object source) {
		pageBuilder.setNull(column);
	}

	public abstract void visitCellValueBoolean(Column column, Object source, boolean value);

	public abstract void visitCellValueError(Column column, Object source, int code);

	public void visitCellFormula(Column column, Cell cell) {
		pageBuilder.setString(column, cell.getCellFormula());
	}

	public abstract void visitValueLong(Column column, Object source, long value);

	public abstract void visitSheetName(Column column);

	public abstract void visitRowNumber(Column column, int index1);

	public abstract void visitColumnNumber(Column column, int index1);
}
