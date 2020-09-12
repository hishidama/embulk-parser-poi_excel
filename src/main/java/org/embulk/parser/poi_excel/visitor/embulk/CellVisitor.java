package org.embulk.parser.poi_excel.visitor.embulk;

import java.text.MessageFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.config.ConfigException;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean.ErrorStrategy;
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

	public abstract void visitSheetName(Column column, Sheet sheet);

	public abstract void visitRowNumber(Column column, int index1);

	public abstract void visitColumnNumber(Column column, int index1);

	protected void doConvertError(Column column, Object srcValue, Throwable t) {
		PoiExcelColumnBean bean = visitorValue.getColumnBean(column);
		ErrorStrategy strategy = bean.getConvertErrorStrategy();
		switch (strategy.getStrategy()) {
		default:
			break;
		case CONSTANT:
			String value = strategy.getValue();
			if (value == null) {
				pageBuilder.setNull(column);
			} else {
				try {
					doConvertErrorConstant(column, value);
				} catch (Exception e) {
					throw new ConfigException(MessageFormat.format("constant value convert error. value={0}", value), e);
				}
			}
			return;
		}

		throw new RuntimeException(MessageFormat.format("convert error. value={0}", srcValue), t);
	}

	protected abstract void doConvertErrorConstant(Column column, String value) throws Exception;
}
