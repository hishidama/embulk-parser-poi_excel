package org.embulk.parser.poi_excel.visitor;

import org.apache.poi.ss.usermodel.Cell;
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

	private static final String[] CELL_TYPE_STRING = { "NUMERIC", "STRING", "FORMULA", "BLANK", "BOOLEAN", "ERROR" };

	public void visit(PoiExcelColumnBean bean, Cell cell, int cellType, CellVisitor visitor) {
		assert cell != null;

		Column column = bean.getColumn();
		if (column.getType() instanceof StringType) {
			String type;
			if (0 <= cellType && cellType < CELL_TYPE_STRING.length) {
				type = CELL_TYPE_STRING[cellType];
			} else {
				type = Integer.toString(cellType);
			}
			visitor.visitCellValueString(column, cell, type);
			return;
		}

		visitor.visitCellValueNumeric(column, cell, cellType);
	}
}
