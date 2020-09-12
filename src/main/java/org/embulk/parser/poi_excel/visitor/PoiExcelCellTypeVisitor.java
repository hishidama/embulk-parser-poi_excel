package org.embulk.parser.poi_excel.visitor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.type.StringType;

public class PoiExcelCellTypeVisitor {
	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public PoiExcelCellTypeVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visit(PoiExcelColumnBean bean, Cell cell, CellType cellType, CellVisitor visitor) {
		assert cell != null;

		Column column = bean.getColumn();
		if (column.getType() instanceof StringType) {
			String type = cellType.name();
			visitor.visitCellValueString(column, cell, type);
			return;
		}

		int code = getCode(cellType);
		visitor.visitCellValueNumeric(column, cell, code);
	}

	private static int getCode(CellType cellType) {
		switch (cellType) {
		case NUMERIC:
			return 0;
		case STRING:
			return 1;
		case FORMULA:
			return 2;
		case BLANK:
			return 3;
		case BOOLEAN:
			return 4;
		case ERROR:
			return 5;
		default:
			return -1;
		}
	}
}
