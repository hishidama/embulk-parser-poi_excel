package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.type.LongType;

public class PoiExcelCellFontVisitor {

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public PoiExcelCellFontVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visitCellFont(Column column, CellStyle style, String name) {
		short index = style.getFontIndex();
		if (name.equals("font") && column.getType() instanceof LongType) {
			pageBuilder.setLong(column, index);
			return;
		}
		Sheet sheet = visitorValue.getSheet();
		Font font = sheet.getWorkbook().getFontAt(index);
		if (font == null) {
			pageBuilder.setNull(column);
			return;
		}
		switch (name) {
		case "font":
			pageBuilder.setString(column, font.toString());
			break;
		case "fontName":
			pageBuilder.setString(column, font.getFontName());
			break;
		case "fontHeight":
			pageBuilder.setLong(column, font.getFontHeight());
			break;
		case "fontHeightInPoints":
			pageBuilder.setLong(column, font.getFontHeightInPoints());
			break;
		case "fontItalic":
			pageBuilder.setBoolean(column, font.getItalic());
			break;
		case "fontStrikeout":
			pageBuilder.setBoolean(column, font.getStrikeout());
			break;
		case "fontColor":
			if (font instanceof XSSFFont) {
				visitCellColor(column, ((XSSFFont) font).getXSSFColor());
			} else {
				visitCellColor(column, font.getColor());
			}
			break;
		case "fontTypeOffset":
			pageBuilder.setLong(column, font.getTypeOffset());
			break;
		case "fontUnderline":
			pageBuilder.setLong(column, font.getUnderline());
			break;
		case "fontCharSet":
			pageBuilder.setLong(column, font.getCharSet());
			break;
		case "fontIndex":
			pageBuilder.setLong(column, font.getIndex());
			break;
		case "fontBoldweight":
			pageBuilder.setLong(column, font.getBoldweight());
			break;
		case "fontBold":
			pageBuilder.setBoolean(column, font.getBold());
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
