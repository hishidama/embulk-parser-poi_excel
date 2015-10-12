package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.type.StringType;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.common.base.Optional;

public class PoiExcelCellFontVisitor {

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public PoiExcelCellFontVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visitCellFont(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
		CellStyle style = cell.getCellStyle();
		short index = style.getFontIndex();
		Workbook book = visitorValue.getSheet().getWorkbook();
		Font font = book.getFontAt(index);
		if (font == null) {
			pageBuilder.setNull(column);
			return;
		}

		String suffix = option.getValueTypeSuffix();
		if (suffix != null) {
			visitCellFontKey(column, option, suffix, cell, font, visitor);
		} else {
			visitCellFontJson(column, option, cell, font, visitor);
		}
	}

	private void visitCellFontKey(Column column, ColumnOptionTask option, String key, Cell cell, Font font,
			CellVisitor visitor) {
		Object value = getFontValue(column, option, cell, font, key);
		if (value == null) {
			pageBuilder.setNull(column);
		} else if (value instanceof String) {
			visitor.visitCellValueString(column, font, (String) value);
		} else if (value instanceof Long) {
			visitor.visitValueLong(column, font, (Long) value);
		} else if (value instanceof Boolean) {
			visitor.visitCellValueBoolean(column, font, (Boolean) value);
		} else {
			throw new IllegalStateException(MessageFormat.format("unsupported conversion. type={0}, value={1}", value
					.getClass().getName(), value));
		}
	}

	private void visitCellFontJson(Column column, ColumnOptionTask option, Cell cell, Font font, CellVisitor visitor) {
		Map<String, Object> result;

		Optional<List<String>> nameOption = option.getCellStyleName();
		if (nameOption.isPresent()) {
			result = new LinkedHashMap<>();

			List<String> list = nameOption.get();
			for (String key : list) {
				Object value = getFontValue(column, option, cell, font, key);
				result.put(key, value);
			}
		} else {
			result = new TreeMap<>();

			for (String key : STYLE_MAP.keySet()) {
				Object value = getFontValue(column, option, cell, font, key);
				result.put(key, value);
			}
		}

		String json;
		try {
			ObjectMapper mapper = new ObjectMapper();
			json = mapper.writeValueAsString(result);
		} catch (JsonProcessingException e) {
			throw new RuntimeException(e);
		}

		visitor.visitCellValueString(column, cell, json);
	}

	protected Object getFontValue(Column column, ColumnOptionTask option, Cell cell, Font font, String key) {
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
		STYLE_MAP = map;
	}
}
