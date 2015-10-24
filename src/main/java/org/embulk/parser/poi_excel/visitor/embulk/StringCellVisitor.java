package org.embulk.parser.poi_excel.visitor.embulk;

import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorValue;
import org.embulk.spi.Column;

public class StringCellVisitor extends CellVisitor {

	public StringCellVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	public void visitCellValueNumeric(Column column, Object source, double value) {
		String s = Double.toString(value);
		if (s.endsWith(".0")) {
			s = s.substring(0, s.length() - 2);
		}
		pageBuilder.setString(column, s);
	}

	@Override
	public void visitCellValueString(Column column, Object source, String value) {
		pageBuilder.setString(column, value);
	}

	@Override
	public void visitCellValueBoolean(Column column, Object source, boolean value) {
		pageBuilder.setString(column, Boolean.toString(value));
	}

	@Override
	public void visitCellValueError(Column column, Object source, int code) {
		FormulaError error = FormulaError.forInt((byte) code);
		String value = error.getString();
		pageBuilder.setString(column, value);
	}

	@Override
	public void visitValueLong(Column column, Object source, long value) {
		String s = Long.toString(value);
		pageBuilder.setString(column, s);
	}

	@Override
	public void visitSheetName(Column column) {
		Sheet sheet = visitorValue.getSheet();
		pageBuilder.setString(column, sheet.getSheetName());
	}

	@Override
	public void visitRowNumber(Column column, int index1) {
		pageBuilder.setString(column, Integer.toString(index1));
	}

	@Override
	public void visitColumnNumber(Column column, int index1) {
		String value = CellReference.convertNumToColString(index1 - 1);
		pageBuilder.setString(column, value);
	}

	@Override
	protected void doConvertErrorConstant(Column column, String value) throws Exception {
		pageBuilder.setString(column, value);
	}
}
