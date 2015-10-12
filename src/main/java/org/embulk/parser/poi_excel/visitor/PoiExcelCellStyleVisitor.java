package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.type.StringType;

import com.google.common.base.Optional;

public class PoiExcelCellStyleVisitor {

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public PoiExcelCellStyleVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visitCellStyle(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
		Optional<String> nameOption = option.getCellStyleName();
		if (!nameOption.isPresent()) {
			throw new RuntimeException(MessageFormat.format("cell_style_name must be specified. column.name={0}",
					column.getName()));
		}

		String name = nameOption.get();
		CellStyle style = cell.getCellStyle();
		switch (name) {
		case CellUtil.ALIGNMENT:
			pageBuilder.setLong(column, style.getAlignment());
			break;
		case CellUtil.BORDER_BOTTOM:
			pageBuilder.setLong(column, style.getBorderBottom());
			break;
		case CellUtil.BORDER_LEFT:
			pageBuilder.setLong(column, style.getBorderLeft());
			break;
		case CellUtil.BORDER_RIGHT:
			pageBuilder.setLong(column, style.getBorderRight());
			break;
		case CellUtil.BORDER_TOP:
			pageBuilder.setLong(column, style.getBorderTop());
			break;
		case "border":
			long border = (style.getBorderTop() << 24) | (style.getBorderBottom() << 16) | (style.getBorderLeft() << 8)
					| style.getBorderRight();
			pageBuilder.setLong(column, border);
			break;
		case CellUtil.BOTTOM_BORDER_COLOR:
			if (style instanceof XSSFCellStyle) {
				visitCellColor(column, ((XSSFCellStyle) style).getBottomBorderXSSFColor());
			} else {
				visitCellColor(column, style.getBottomBorderColor());
			}
			break;
		case CellUtil.DATA_FORMAT:
			if (column.getType() instanceof StringType) {
				pageBuilder.setString(column, style.getDataFormatString());
			} else {
				pageBuilder.setLong(column, style.getDataFormat());
			}
			break;
		case CellUtil.FILL_BACKGROUND_COLOR:
			visitCellColor(column, style.getFillBackgroundColorColor());
			break;
		case CellUtil.FILL_FOREGROUND_COLOR:
			visitCellColor(column, style.getFillForegroundColorColor());
			break;
		case CellUtil.FILL_PATTERN:
			pageBuilder.setLong(column, style.getFillPattern());
			break;
		case CellUtil.HIDDEN:
			pageBuilder.setBoolean(column, style.getHidden());
			break;
		case CellUtil.INDENTION:
			pageBuilder.setLong(column, style.getIndention());
			break;
		case CellUtil.LEFT_BORDER_COLOR:
			if (style instanceof XSSFCellStyle) {
				visitCellColor(column, ((XSSFCellStyle) style).getLeftBorderXSSFColor());
			} else {
				visitCellColor(column, style.getLeftBorderColor());
			}
			break;
		case CellUtil.LOCKED:
			pageBuilder.setBoolean(column, style.getLocked());
			break;
		case CellUtil.RIGHT_BORDER_COLOR:
			if (style instanceof XSSFCellStyle) {
				visitCellColor(column, ((XSSFCellStyle) style).getRightBorderXSSFColor());
			} else {
				visitCellColor(column, style.getRightBorderColor());
			}
			break;
		case CellUtil.ROTATION:
			pageBuilder.setLong(column, style.getRotation());
			break;
		case CellUtil.TOP_BORDER_COLOR:
			if (style instanceof XSSFCellStyle) {
				visitCellColor(column, ((XSSFCellStyle) style).getTopBorderXSSFColor());
			} else {
				visitCellColor(column, style.getTopBorderColor());
			}
			break;
		case CellUtil.VERTICAL_ALIGNMENT:
			pageBuilder.setLong(column, style.getVerticalAlignment());
			break;
		case CellUtil.WRAP_TEXT:
			pageBuilder.setBoolean(column, style.getWrapText());
			break;
		default:
			throw new UnsupportedOperationException(MessageFormat.format("unsupported cell_style_name={0}", name));
		}
	}

	protected final void visitCellColor(Column column, short colorIndex) {
		PoiExcelVisitorFactory factory = visitorValue.getVisitorFactory();
		PoiExcelColorVisitor delegator = factory.getPoiExcelColorVisitor();
		delegator.visitCellColor(column, colorIndex);
	}

	protected final void visitCellColor(Column column, Color color) {
		PoiExcelVisitorFactory factory = visitorValue.getVisitorFactory();
		PoiExcelColorVisitor delegator = factory.getPoiExcelColorVisitor();
		delegator.visitCellColor(column, color);
	}
}
