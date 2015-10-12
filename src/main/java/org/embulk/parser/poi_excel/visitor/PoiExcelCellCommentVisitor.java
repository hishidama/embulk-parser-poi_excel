package org.embulk.parser.poi_excel.visitor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;

public class PoiExcelCellCommentVisitor {

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public PoiExcelCellCommentVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visitCellComment(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
		Comment comment = cell.getCellComment();
		if (comment == null) {
			pageBuilder.setNull(column);
			return;
		}
		String s = comment.getString().getString();
		pageBuilder.setString(column, s);
	}
}
