package org.embulk.parser.poi_excel.visitor.embulk;

import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorValue;
import org.embulk.spi.Column;

public class LongCellVisitor extends CellVisitor {

	public LongCellVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	public void visitCellValueNumeric(Column column, Object source, double value) {
		pageBuilder.setLong(column, (long) value);
	}

	@Override
	public void visitCellValueString(Column column, Object source, String value) {
		long l;
		try {
			l = Long.parseLong(value);
		} catch (NumberFormatException e) {
			doConvertError(column, value, e);
			return;
		}
		pageBuilder.setLong(column, l);
	}

	@Override
	public void visitCellValueBoolean(Column column, Object source, boolean value) {
		pageBuilder.setLong(column, value ? 1 : 0);
	}

	@Override
	public void visitCellValueError(Column column, Object source, int code) {
		pageBuilder.setLong(column, code);
	}

	@Override
	public void visitValueLong(Column column, Object source, long value) {
		pageBuilder.setLong(column, value);
	}

	@Override
	public void visitSheetName(Column column) {
		Sheet sheet = visitorValue.getSheet();
		int index = sheet.getWorkbook().getSheetIndex(sheet);
		pageBuilder.setLong(column, index);
	}

	@Override
	public void visitRowNumber(Column column, int index1) {
		pageBuilder.setLong(column, index1);
	}

	@Override
	public void visitColumnNumber(Column column, int index1) {
		pageBuilder.setLong(column, index1);
	}

	@Override
	protected void doConvertErrorConstant(Column column, String value) throws Exception {
		pageBuilder.setLong(column, Long.parseLong(value));
	}
}
