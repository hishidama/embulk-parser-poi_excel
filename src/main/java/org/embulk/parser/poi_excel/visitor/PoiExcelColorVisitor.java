package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.type.StringType;

public class PoiExcelColorVisitor {

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public PoiExcelColorVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visitCellColor(Column column, short colorIndex) {
		HSSFWorkbook book = (HSSFWorkbook) visitorValue.getSheet().getWorkbook();
		HSSFPalette palette = book.getCustomPalette();
		HSSFColor color = palette.getColor(colorIndex);
		visitCellColor(column, color);
	}

	public void visitCellColor(Column column, Color color) {
		int rgb = getRGB(color);
		if (rgb < 0) {
			pageBuilder.setNull(column);
			return;
		}

		if (column.getType() instanceof StringType) {
			String s = String.format("%06x", rgb);
			pageBuilder.setString(column, s);
		} else {
			pageBuilder.setLong(column, rgb);
		}
	}

	public static Color getColor(Workbook workbook, short colorIndex) {
		HSSFWorkbook book = (HSSFWorkbook) workbook;
		HSSFPalette palette = book.getCustomPalette();
		HSSFColor color = palette.getColor(colorIndex);
		return color;
	}

	public static int getRGB(Color color) {
		if (color == null) {
			return -1;
		}

		int[] rgb = new int[3];
		if (color instanceof HSSFColor) {
			HSSFColor hssf = (HSSFColor) color;
			short[] s = hssf.getTriplet();
			rgb[0] = s[0] & 0xff;
			rgb[1] = s[1] & 0xff;
			rgb[2] = s[2] & 0xff;
		} else if (color instanceof XSSFColor) {
			XSSFColor xssf = (XSSFColor) color;
			byte[] b = xssf.getRGB();
			rgb[0] = b[0] & 0xff;
			rgb[1] = b[1] & 0xff;
			rgb[2] = b[2] & 0xff;
		} else {
			throw new IllegalStateException(MessageFormat.format("unsupported POI color={0}", color));
		}

		return (rgb[0] << 16) | (rgb[1] << 8) | rgb[2];
	}
}
