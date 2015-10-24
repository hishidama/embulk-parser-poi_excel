package org.embulk.parser.poi_excel.visitor.embulk;

import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorValue;
import org.embulk.spi.Column;

public class BooleanCellVisitor extends CellVisitor {

	public BooleanCellVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	public void visitCellValueNumeric(Column column, Object source, double value) {
		pageBuilder.setBoolean(column, value != 0d);
	}

	@Override
	public void visitCellValueString(Column column, Object source, String value) {
		pageBuilder.setBoolean(column, Boolean.parseBoolean(value));
	}

	@Override
	public void visitCellValueBoolean(Column column, Object source, boolean value) {
		pageBuilder.setBoolean(column, value);
	}

	@Override
	public void visitCellValueError(Column column, Object source, int code) {
		pageBuilder.setNull(column);
	}

	@Override
	public void visitValueLong(Column column, Object source, long value) {
		pageBuilder.setBoolean(column, value != 0);
	}

	@Override
	public void visitSheetName(Column column) {
		Sheet sheet = visitorValue.getSheet();
		int index = sheet.getWorkbook().getSheetIndex(sheet);
		pageBuilder.setBoolean(column, index != 0);
	}

	@Override
	public void visitRowNumber(Column column, int index1) {
		pageBuilder.setBoolean(column, index1 != 0);
	}

	@Override
	public void visitColumnNumber(Column column, int index1) {
		pageBuilder.setBoolean(column, index1 != 0);
	}

	@Override
	protected void doConvertErrorConstant(Column column, String value) throws Exception {
		pageBuilder.setBoolean(column, Boolean.parseBoolean(value));
	}
}
