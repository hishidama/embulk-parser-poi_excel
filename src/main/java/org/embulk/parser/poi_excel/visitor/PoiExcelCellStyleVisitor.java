package org.embulk.parser.poi_excel.visitor;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.spi.Column;
import org.embulk.spi.type.StringType;

public class PoiExcelCellStyleVisitor extends AbstractPoiExcelCellAttributeVisitor<CellStyle> {

	public PoiExcelCellStyleVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	protected CellStyle getAttributeSource(PoiExcelColumnBean bean, Cell cell) {
		return cell.getCellStyle();
	}

	protected boolean acceptKey(String key) {
		if (key.equals("border")) {
			return false;
		}
		return true;
	}

	@Override
	protected Map<String, AttributeSupplier<CellStyle>> getAttributeSupplierMap() {
		return SUPPLIER_MAP;
	}

	private static final Map<String, AttributeSupplier<CellStyle>> SUPPLIER_MAP;
	static {
		Map<String, AttributeSupplier<CellStyle>> map = new HashMap<>(32);
		map.put("alignment", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getAlignment().getCode();
			}
		});
		map.put("border", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				int n0 = style.getBorderTop().getCode();
				int n1 = style.getBorderBottom().getCode();
				int n2 = style.getBorderLeft().getCode();
				int n3 = style.getBorderRight().getCode();
				if (column.getType() instanceof StringType) {
					return String.format("%02x%02x%02x%02x", n0, n1, n2, n3);
				}
				return (long) ((n0 << 24) | (n1 << 16) | (n2 << 8) | n3);
			}
		});
		map.put("border_bottom", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getBorderBottom().getCode();
			}
		});
		map.put("border_left", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getBorderLeft().getCode();
			}
		});
		map.put("border_right", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getBorderRight().getCode();
			}
		});
		map.put("border_top", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getBorderTop().getCode();
			}
		});
		map.put("border_bottom_color", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				if (style instanceof XSSFCellStyle) {
					return ((XSSFCellStyle) style).getBottomBorderXSSFColor();
				} else {
					Workbook book = cell.getSheet().getWorkbook();
					short color = style.getBottomBorderColor();
					return PoiExcelColorVisitor.getHssfColor(book, color);
				}
			}
		});
		map.put("border_left_color", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				if (style instanceof XSSFCellStyle) {
					return ((XSSFCellStyle) style).getLeftBorderXSSFColor();
				} else {
					Workbook book = cell.getSheet().getWorkbook();
					short color = style.getLeftBorderColor();
					return PoiExcelColorVisitor.getHssfColor(book, color);
				}
			}
		});
		map.put("border_right_color", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				if (style instanceof XSSFCellStyle) {
					return ((XSSFCellStyle) style).getRightBorderXSSFColor();
				} else {
					Workbook book = cell.getSheet().getWorkbook();
					short color = style.getRightBorderColor();
					return PoiExcelColorVisitor.getHssfColor(book, color);
				}
			}
		});
		map.put("border_top_color", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				if (style instanceof XSSFCellStyle) {
					return ((XSSFCellStyle) style).getTopBorderXSSFColor();
				} else {
					Workbook book = cell.getSheet().getWorkbook();
					short color = style.getTopBorderColor();
					return PoiExcelColorVisitor.getHssfColor(book, color);
				}
			}
		});
		map.put("data_format", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				if (column.getType() instanceof StringType) {
					return style.getDataFormatString();
				} else {
					return (long) style.getDataFormat();
				}
			}
		});
		map.put("fill_background_color", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return style.getFillBackgroundColorColor();
			}
		});
		map.put("fill_foreground_color", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return style.getFillForegroundColorColor();
			}
		});
		map.put("fill_pattern", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getFillPattern().getCode();
			}
		});
		map.put("font_index", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getFontIndex();
			}
		});
		map.put("hidden", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return style.getHidden();
			}
		});
		map.put("indention", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getIndention();
			}
		});
		map.put("locked", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return style.getLocked();
			}
		});
		map.put("rotation", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getRotation();
			}
		});
		map.put("vertical_alignment", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return (long) style.getVerticalAlignment().getCode();
			}
		});
		map.put("wrap_text", new AttributeSupplier<CellStyle>() {
			@Override
			public Object get(Column column, Cell cell, CellStyle style) {
				return style.getWrapText();
			}
		});
		SUPPLIER_MAP = Collections.unmodifiableMap(map);
	}
}
