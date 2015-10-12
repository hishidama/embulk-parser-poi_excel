package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.spi.Column;
import org.embulk.spi.type.StringType;

public class PoiExcelCellFontVisitor extends AbstractPoiExcelCellAttributeVisitor<Font> {

	public PoiExcelCellFontVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	protected Font getAttributeSource(Column column, ColumnOptionTask option, Cell cell) {
		CellStyle style = cell.getCellStyle();
		short index = style.getFontIndex();
		Workbook book = visitorValue.getSheet().getWorkbook();
		return book.getFontAt(index);
	}

	@Override
	protected Collection<String> geyAllKeys() {
		return STYLE_MAP.keySet();
	}

	@Override
	protected Object getAttributeValue(Column column, ColumnOptionTask option, Cell cell, Font font, String key) {
		CellStyleSupplier supplier = STYLE_MAP.get(key.toLowerCase());
		if (supplier == null) {
			throw new UnsupportedOperationException(MessageFormat.format(
					"unsupported cell font name={0}, choose in {1}", key, new TreeSet<>(STYLE_MAP.keySet())));
		}
		Object value = supplier.get(column, cell, font);
		if (value instanceof Color) {
			int rgb = PoiExcelColorVisitor.getRGB((Color) value);
			if (column.getType() instanceof StringType) {
				value = String.format("%06x", rgb);
			} else {
				value = (long) rgb;
			}
		}
		return value;
	}

	protected static interface CellStyleSupplier {
		public Object get(Column column, Cell cell, Font font);
	}

	protected static final Map<String, CellStyleSupplier> STYLE_MAP;
	static {
		Map<String, CellStyleSupplier> map = new HashMap<>(32);
		map.put("font_name", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return font.getFontName();
			}
		});
		map.put("font_height", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getFontHeight();
			}
		});
		map.put("font_height_in_points", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getFontHeightInPoints();
			}
		});
		map.put("italic", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return font.getItalic();
			}
		});
		map.put("strikeout", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return font.getStrikeout();
			}
		});
		map.put("color", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				if (font instanceof XSSFFont) {
					return ((XSSFFont) font).getXSSFColor();
				} else {
					Workbook book = cell.getSheet().getWorkbook();
					short color = font.getColor();
					return PoiExcelColorVisitor.getColor(book, color);
				}
			}
		});
		map.put("type_offset", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getTypeOffset();
			}
		});
		map.put("underline", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getUnderline();
			}
		});
		map.put("char_set", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getCharSet();
			}
		});
		map.put("index", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getIndex();
			}
		});
		map.put("boldweight", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return (long) font.getBoldweight();
			}
		});
		map.put("bold", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, Font font) {
				return font.getBold();
			}
		});
		STYLE_MAP = Collections.unmodifiableMap(map);
	}
}
