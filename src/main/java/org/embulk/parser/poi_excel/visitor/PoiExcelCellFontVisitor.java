package org.embulk.parser.poi_excel.visitor;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.spi.Column;

public class PoiExcelCellFontVisitor extends AbstractPoiExcelCellAttributeVisitor<Font> {

	public PoiExcelCellFontVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	protected Font getAttributeSource(PoiExcelColumnBean bean, Cell cell) {
		CellStyle style = cell.getCellStyle();
		int index = style.getFontIndex();
		Workbook book = visitorValue.getSheet().getWorkbook();
		return book.getFontAt(index);
	}

	@Override
	protected Map<String, AttributeSupplier<Font>> getAttributeSupplierMap() {
		return SUPPLIER_MAP;
	}

	private static final Map<String, AttributeSupplier<Font>> SUPPLIER_MAP;
	static {
		Map<String, AttributeSupplier<Font>> map = new HashMap<>(32);
		map.put("font_name", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return font.getFontName();
			}
		});
		map.put("font_height", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getFontHeight();
			}
		});
		map.put("font_height_in_points", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getFontHeightInPoints();
			}
		});
		map.put("italic", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return font.getItalic();
			}
		});
		map.put("strikeout", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return font.getStrikeout();
			}
		});
		map.put("color", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				if (font instanceof XSSFFont) {
					return ((XSSFFont) font).getXSSFColor();
				} else {
					Workbook book = cell.getSheet().getWorkbook();
					short color = font.getColor();
					return PoiExcelColorVisitor.getHssfColor(book, color);
				}
			}
		});
		map.put("type_offset", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getTypeOffset();
			}
		});
		map.put("underline", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getUnderline();
			}
		});
		map.put("char_set", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getCharSet();
			}
		});
		map.put("index", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getIndex();
			}
		});
		map.put("bold", new AttributeSupplier<Font>() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return font.getBold();
			}
		});
		SUPPLIER_MAP = Collections.unmodifiableMap(map);
	}
}
