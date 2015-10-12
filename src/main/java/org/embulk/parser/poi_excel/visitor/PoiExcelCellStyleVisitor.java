package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.spi.Column;
import org.embulk.spi.type.StringType;

public class PoiExcelCellStyleVisitor extends AbstractPoiExcelCellAttributeVisitor<CellStyle> {

	public PoiExcelCellStyleVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	protected CellStyle getAttributeSource(Column column, ColumnOptionTask option, Cell cell) {
		return cell.getCellStyle();
	}

	@Override
	protected Collection<String> geyAllKeys() {
		return ALL_KEYS;
	}

	@Override
	protected Object getAttributeValue(Column column, ColumnOptionTask option, Cell cell, CellStyle style, String key) {
		CellStyleSupplier supplier = STYLE_MAP.get(key.toLowerCase());
		if (supplier == null) {
			throw new UnsupportedOperationException(MessageFormat.format(
					"unsupported cell style name={0}, choose in {1}", key, new TreeSet<>(STYLE_MAP.keySet())));
		}
		Object value = supplier.get(column, cell, style);
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
		public Object get(Column column, Cell cell, CellStyle style);
	}

	protected static final Map<String, CellStyleSupplier> STYLE_MAP;
	protected static final Set<String> ALL_KEYS;

	static {
		Map<String, CellStyleSupplier> map = new HashMap<>(32);
		map.put("alignment", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getAlignment();
			}
		});
		map.put("border", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				int n0 = style.getBorderTop();
				int n1 = style.getBorderBottom();
				int n2 = style.getBorderLeft();
				int n3 = style.getBorderRight();
				if (column.getType() instanceof StringType) {
					return String.format("%02x%02x%02x%02x", n0, n1, n2, n3);
				}
				return (long) ((n0 << 24) | (n1 << 16) | (n2 << 8) | n3);
			}
		});
		map.put("border_bottom", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getBorderBottom();
			}
		});
		map.put("border_left", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getBorderLeft();
			}
		});
		map.put("border_right", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getBorderRight();
			}
		});
		map.put("border_top", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getBorderTop();
			}
		});
		map.put("border_bottom_color", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				if (style instanceof XSSFCellStyle) {
					return ((XSSFCellStyle) style).getBottomBorderXSSFColor();
				} else {
					Workbook book = cell.getSheet().getWorkbook();
					short color = style.getBottomBorderColor();
					return PoiExcelColorVisitor.getColor(book, color);
				}
			}
		});
		map.put("border_left_color", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				if (style instanceof XSSFCellStyle) {
					return ((XSSFCellStyle) style).getLeftBorderXSSFColor();
				} else {
					Workbook book = cell.getSheet().getWorkbook();
					short color = style.getLeftBorderColor();
					return PoiExcelColorVisitor.getColor(book, color);
				}
			}
		});
		map.put("border_right_color", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				if (style instanceof XSSFCellStyle) {
					return ((XSSFCellStyle) style).getRightBorderXSSFColor();
				} else {
					Workbook book = cell.getSheet().getWorkbook();
					short color = style.getRightBorderColor();
					return PoiExcelColorVisitor.getColor(book, color);
				}
			}
		});
		map.put("border_top_color", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				if (style instanceof XSSFCellStyle) {
					return ((XSSFCellStyle) style).getTopBorderXSSFColor();
				} else {
					Workbook book = cell.getSheet().getWorkbook();
					short color = style.getTopBorderColor();
					return PoiExcelColorVisitor.getColor(book, color);
				}
			}
		});
		map.put("data_format", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				if (column.getType() instanceof StringType) {
					return style.getDataFormatString();
				} else {
					return (long) style.getDataFormat();
				}
			}
		});
		map.put("fill_background_color", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return style.getFillBackgroundColorColor();
			}
		});
		map.put("fill_foreground_color", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return style.getFillForegroundColorColor();
			}
		});
		map.put("fill_pattern", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getFillPattern();
			}
		});
		map.put("font_index", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getFontIndex();
			}
		});
		map.put("hidden", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return style.getHidden();
			}
		});
		map.put("indention", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getIndention();
			}
		});
		map.put("locked", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return style.getLocked();
			}
		});
		map.put("rotation", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getRotation();
			}
		});
		map.put("vertical_alignment", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getVerticalAlignment();
			}
		});
		map.put("wrap_text", new CellStyleSupplier() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return style.getWrapText();
			}
		});
		STYLE_MAP = Collections.unmodifiableMap(map);

		Set<String> set = new HashSet<String>(STYLE_MAP.keySet());
		set.remove("border");
		ALL_KEYS = Collections.unmodifiableSet(set);
	}
}
