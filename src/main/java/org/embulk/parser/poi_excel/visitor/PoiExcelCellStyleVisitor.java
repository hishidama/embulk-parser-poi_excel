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
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.type.StringType;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.common.base.Optional;

public class PoiExcelCellStyleVisitor {

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public PoiExcelCellStyleVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visitCellStyle(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
		CellStyle style = cell.getCellStyle();

		String suffix = option.getValueTypeSuffix();
		if (suffix != null) {
			visitCellStyleKey(column, option, suffix, cell, style, visitor);
		} else {
			visitCellStyleJson(column, option, cell, style, visitor);
		}
	}

	private void visitCellStyleKey(Column column, ColumnOptionTask option, String key, Cell cell, CellStyle style,
			CellVisitor visitor) {
		Object value = getStyleValue(column, option, cell, style, key);
		if (value == null) {
			pageBuilder.setNull(column);
		} else if (value instanceof String) {
			visitor.visitCellValueString(column, style, (String) value);
		} else if (value instanceof Long) {
			visitor.visitValueLong(column, style, (Long) value);
		} else if (value instanceof Boolean) {
			visitor.visitCellValueBoolean(column, style, (Boolean) value);
		} else {
			throw new IllegalStateException(MessageFormat.format("unsupported conversion. type={0}, value={1}", value
					.getClass().getName(), value));
		}
	}

	private void visitCellStyleJson(Column column, ColumnOptionTask option, Cell cell, CellStyle style,
			CellVisitor visitor) {
		Map<String, Object> result;

		Optional<List<String>> nameOption = option.getCellStyleName();
		if (nameOption.isPresent()) {
			result = new LinkedHashMap<>();

			List<String> list = nameOption.get();
			for (String key : list) {
				Object value = getStyleValue(column, option, cell, style, key);
				result.put(key, value);
			}
		} else {
			result = new TreeMap<>();

			for (String key : STYLE_MAP.keySet()) {
				if (key.equals("border")) {
					continue;
				}

				Object value = getStyleValue(column, option, cell, style, key);
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

	protected Object getStyleValue(Column column, ColumnOptionTask option, Cell cell, CellStyle style, String key) {
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
		STYLE_MAP = map;
	}
}
