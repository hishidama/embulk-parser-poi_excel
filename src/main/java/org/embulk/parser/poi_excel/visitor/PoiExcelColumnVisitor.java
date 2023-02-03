package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.parser.poi_excel.PoiExcelColumnValueType;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.parser.poi_excel.bean.record.PoiExcelRecord;
import org.embulk.parser.poi_excel.bean.util.PoiExcelCellAddress;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.ColumnVisitor;
import org.embulk.spi.Exec;
import org.embulk.spi.PageBuilder;
import org.slf4j.Logger;

public class PoiExcelColumnVisitor implements ColumnVisitor {
	private final Logger log = Exec.getLogger(getClass());

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;
	protected final PoiExcelVisitorFactory factory;

	protected PoiExcelRecord record;

	public PoiExcelColumnVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
		this.factory = visitorValue.getVisitorFactory();
	}

	public void setRecord(PoiExcelRecord record) {
		this.record = record;
	}

	@Override
	public final void booleanColumn(Column column) {
		visitCell0(column, factory.getBooleanCellVisitor());
	}

	@Override
	public final void longColumn(Column column) {
		visitCell0(column, factory.getLongCellVisitor());
	}

	@Override
	public final void doubleColumn(Column column) {
		visitCell0(column, factory.getDoubleCellVisitor());
	}

	@Override
	public final void stringColumn(Column column) {
		visitCell0(column, factory.getStringCellVisitor());
	}

	@Override
	public final void timestampColumn(Column column) {
		visitCell0(column, factory.getTimestampCellVisitor());
	}

	@Override
	public final void jsonColumn(Column column) {
		visitCell0(column, factory.getStringCellVisitor());
	}

		protected final void visitCell0(Column column, CellVisitor visitor) {
		if (log.isTraceEnabled()) {
			log.trace("{} start", column);
		}
		try {
			visitCell(column, visitor);
		} catch (Exception e) {
			String sheetName = visitorValue.getSheet().getSheetName();
			String ref = record.getCellReference(visitorValue.getColumnBean(column)).formatAsString();
			throw new RuntimeException(MessageFormat.format("error at {0} cell={1}!{2}. {3}", column, sheetName, ref,
					e.getMessage()), e);
		}
		if (log.isTraceEnabled()) {
			log.trace("{} end", column);
		}
	}

	protected void visitCell(Column column, CellVisitor visitor) {
		PoiExcelColumnBean bean = visitorValue.getColumnBean(column);
		PoiExcelColumnValueType valueType = bean.getValueType();
		PoiExcelCellAddress cellAddress = bean.getCellAddress();

		switch (valueType) {
		case SHEET_NAME:
			if (cellAddress != null) {
				Sheet sheet = cellAddress.getSheet(record);
				visitor.visitSheetName(column, sheet);
			} else {
				visitor.visitSheetName(column);
			}
			return;
		case ROW_NUMBER:
			int rowIndex;
			if (cellAddress != null) {
				rowIndex = cellAddress.getRowIndex();
			} else {
				rowIndex = record.getRowIndex(bean);
			}
			visitor.visitRowNumber(column, rowIndex + 1);
			return;
		case COLUMN_NUMBER:
			int columnIndex;
			if (cellAddress != null) {
				columnIndex = cellAddress.getColumnIndex();
			} else {
				columnIndex = record.getColumnIndex(bean);
			}
			visitor.visitColumnNumber(column, columnIndex + 1);
			return;
		case CONSTANT:
			visitCellConstant(column, bean.getValueTypeSuffix(), visitor);
			return;
		default:
			break;
		}

		// assert valueType.useCell();
		Cell cell;
		if (cellAddress != null) {
			cell = cellAddress.getCell(record);
		} else {
			cell = record.getCell(bean);
		}
		if (cell == null) {
			visitCellNull(column);
			return;
		}
		switch (valueType) {
		case CELL_VALUE:
		case CELL_FORMULA:
			visitCellValue(bean, cell, visitor);
			return;
		case CELL_STYLE:
			visitCellStyle(bean, cell, visitor);
			return;
		case CELL_FONT:
			visitCellFont(bean, cell, visitor);
			return;
		case CELL_COMMENT:
			visitCellComment(bean, cell, visitor);
			return;
		case CELL_TYPE:
			visitCellType(bean, cell, cell.getCellTypeEnum(), visitor);
			return;
		case CELL_CACHED_TYPE:
			if (cell.getCellTypeEnum() == CellType.FORMULA) {
				visitCellType(bean, cell, cell.getCachedFormulaResultTypeEnum(), visitor);
			} else {
				visitCellType(bean, cell, cell.getCellTypeEnum(), visitor);
			}
			return;
		default:
			throw new UnsupportedOperationException(MessageFormat.format("unsupported value_type={0}", valueType));
		}
	}

	protected void visitCellConstant(Column column, String value, CellVisitor visitor) {
		if (value == null) {
			pageBuilder.setNull(column);
			return;
		}
		visitor.visitCellValueString(column, null, value);
	}

	protected void visitCellNull(Column column) {
		pageBuilder.setNull(column);
	}

	private void visitCellValue(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
		PoiExcelCellValueVisitor delegator = factory.getPoiExcelCellValueVisitor();
		delegator.visitCellValue(bean, cell, visitor);
	}

	private void visitCellStyle(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
		PoiExcelCellStyleVisitor delegator = factory.getPoiExcelCellStyleVisitor();
		delegator.visit(bean, cell, visitor);
	}

	private void visitCellFont(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
		PoiExcelCellFontVisitor delegator = factory.getPoiExcelCellFontVisitor();
		delegator.visit(bean, cell, visitor);
	}

	private void visitCellComment(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
		PoiExcelCellCommentVisitor delegator = factory.getPoiExcelCellCommentVisitor();
		delegator.visit(bean, cell, visitor);
	}

	private void visitCellType(PoiExcelColumnBean bean, Cell cell, CellType cellType, CellVisitor visitor) {
		PoiExcelCellTypeVisitor delegator = factory.getPoiExcelCellTypeVisitor();
		delegator.visit(bean, cell, cellType, visitor);
	}
}
