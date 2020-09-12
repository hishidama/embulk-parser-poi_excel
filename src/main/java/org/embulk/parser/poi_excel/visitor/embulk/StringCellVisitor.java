package org.embulk.parser.poi_excel.visitor.embulk;

import java.text.MessageFormat;

import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorValue;
import org.embulk.spi.Column;

public class StringCellVisitor extends CellVisitor {

	public StringCellVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	public void visitCellValueNumeric(Column column, Object source, double value) {
		String s = toString(column, source, value);
		pageBuilder.setString(column, s);
	}

	protected String toString(Column column, Object source, double value) {
		String format = getNumericFormat(column);
		if (!format.isEmpty()) {
			try {
				return String.format(format, value);
			} catch (Exception e) {
				throw new IllegalArgumentException(MessageFormat.format(
						"illegal String.format for double. numeric_format=\"{0}\"", format), e);
			}
		}

		String s = Double.toString(value);
		if (s.endsWith(".0")) {
			return s.substring(0, s.length() - 2);
		}
		return s;
	}

	protected String getNumericFormat(Column column) {
		PoiExcelColumnBean bean = visitorValue.getColumnBean(column);
		return bean.getNumericFormat();
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
		visitSheetName(column, sheet);
	}

	@Override
	public void visitSheetName(Column column, Sheet sheet) {
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
