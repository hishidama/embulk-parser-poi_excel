package org.embulk.parser.poi_excel;

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.FormulaReplaceTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.spi.Column;
import org.embulk.spi.ColumnConfig;
import org.embulk.spi.ColumnVisitor;
import org.embulk.spi.Exec;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.Schema;
import org.embulk.spi.SchemaConfig;
import org.embulk.spi.time.Timestamp;
import org.embulk.spi.time.TimestampParser;
import org.embulk.spi.type.LongType;
import org.embulk.spi.type.StringType;
import org.embulk.spi.util.Timestamps;
import org.slf4j.Logger;

import com.google.common.base.Optional;

public class PoiExcelColumnVisitor implements ColumnVisitor {
	private final Logger log = Exec.getLogger(getClass());

	protected final PluginTask task;
	protected final Sheet sheet;
	protected final PageBuilder pageBuilder;

	protected Row currentRow;

	public PoiExcelColumnVisitor(PluginTask task, Sheet sheet, PageBuilder pageBuilder) {
		this.task = task;
		this.sheet = sheet;
		this.pageBuilder = pageBuilder;

		initializeColumnOptions();
	}

	protected void initializeColumnOptions() {
		int index = -1;

		Schema schema = task.getColumns().toSchema();
		for (Column column : schema.getColumns()) {
			ColumnOptionTask option = getColumnOption(column);

			PoiExcelColumnValueType valueType;
			if (option.getCellStyleName().isPresent()) {
				valueType = PoiExcelColumnValueType.CELL_STYLE;
			} else {
				String type = option.getValueType();
				try {
					valueType = PoiExcelColumnValueType.valueOf(type.toUpperCase());
				} catch (Exception e) {
					throw new RuntimeException(MessageFormat.format("illegal value_type={0}", type));
				}
			}
			option.setValueTypeEnum(valueType);

			if (valueType.useCell()) {
				Optional<String> numberOption = option.getColumnNumber();
				if (numberOption.isPresent()) {
					String s = numberOption.get();
					index = convertColumnIndex(s);
					option.setColumnIndex(index);
				} else {
					if (valueType.nextIndex()) {
						index++;
					}
					if (index < 0) {
						index = 0;
					}
					option.setColumnIndex(index);
				}
			}
		}
	}

	protected int convertColumnIndex(String s) {
		int index;
		try {
			char c = s.charAt(0);
			if ('0' <= c && c <= '9') {
				index = Integer.parseInt(s) - 1;
			} else {
				index = CellReference.convertColStringToIndex(s);
			}
		} catch (Exception e) {
			throw new RuntimeException(MessageFormat.format("illegal column_index={0}", s), e);
		}
		if (index < 0) {
			throw new RuntimeException(MessageFormat.format("illegal column_index={0}", s));
		}
		return index;
	}

	public void setRow(Row row) {
		this.currentRow = row;
	}

	private List<ColumnOptionTask> columnOptions;

	protected ColumnOptionTask getColumnOption(Column column) {
		if (columnOptions == null) {
			SchemaConfig schemaConfig = task.getColumns();
			columnOptions = new ArrayList<>(schemaConfig.getColumnCount());
			for (ColumnConfig c : schemaConfig.getColumns()) {
				ColumnOptionTask option = c.getOption().loadConfig(ColumnOptionTask.class);
				columnOptions.add(option);
			}
		}
		return columnOptions.get(column.getIndex());
	}

	private TimestampParser[] timestampParsers;

	protected TimestampParser getTimestampParser(Column column) {
		if (timestampParsers == null) {
			timestampParsers = Timestamps.newTimestampColumnParsers(task, task.getColumns());
		}
		return timestampParsers[column.getIndex()];
	}

	private CellVisitor booleanVisitor;

	@Override
	public final void booleanColumn(Column column) {
		if (booleanVisitor == null) {
			booleanVisitor = newBooleanCellVisitor();
		}
		visitCell0(column, booleanVisitor);
	}

	protected CellVisitor newBooleanCellVisitor() {
		return new BooleanCellVisitor();
	}

	private CellVisitor longVisitor;

	@Override
	public final void longColumn(Column column) {
		if (longVisitor == null) {
			longVisitor = newLongCellVisitor();
		}
		visitCell0(column, longVisitor);
	}

	protected CellVisitor newLongCellVisitor() {
		return new LongCellVisitor();
	}

	private CellVisitor doubleVisitor;

	@Override
	public final void doubleColumn(Column column) {
		if (doubleVisitor == null) {
			doubleVisitor = newDoubleCellVisitor();
		}
		visitCell0(column, doubleVisitor);
	}

	protected CellVisitor newDoubleCellVisitor() {
		return new DoubleCellVisitor();
	}

	private CellVisitor stringVisitor;

	@Override
	public final void stringColumn(Column column) {
		if (stringVisitor == null) {
			stringVisitor = newStringCellVisitor();
		}
		visitCell0(column, stringVisitor);
	}

	protected CellVisitor newStringCellVisitor() {
		return new StringCellVisitor();
	}

	private CellVisitor timestampVisitor;

	@Override
	public final void timestampColumn(Column column) {
		if (timestampVisitor == null) {
			timestampVisitor = newTimestampCellVisitor();
		}
		visitCell0(column, timestampVisitor);
	}

	protected CellVisitor newTimestampCellVisitor() {
		return new TimestampCellVisitor();
	}

	protected abstract class CellVisitor {

		public abstract void visitCellValueNumeric(Column column, Object cell, double value);

		public abstract void visitCellValueString(Column column, Object cell, String value);

		public void visitCellValueBlank(Column column, Object cell) {
			pageBuilder.setNull(column);
		}

		public abstract void visitCellValueBoolean(Column column, Object cell, boolean value);

		public abstract void visitCellValueError(Column column, Object cell, int code);

		public void visitCellFormula(Column column, Cell cell) {
			pageBuilder.setString(column, cell.getCellFormula());
		}

		public void visitSheetName(Column column, String sheetName) {
			pageBuilder.setString(column, sheetName);
		}

		public abstract void visitRowNumber(Column column, int index1);

		public abstract void visitColumnNumber(Column column, int index1);
	}

	protected class BooleanCellVisitor extends CellVisitor {

		@Override
		public void visitCellValueNumeric(Column column, Object cell, double value) {
			pageBuilder.setBoolean(column, value != 0d);
		}

		@Override
		public void visitCellValueString(Column column, Object cell, String value) {
			pageBuilder.setBoolean(column, Boolean.parseBoolean(value));
		}

		@Override
		public void visitCellValueBoolean(Column column, Object cell, boolean value) {
			pageBuilder.setBoolean(column, value);
		}

		@Override
		public void visitCellValueError(Column column, Object cell, int code) {
			pageBuilder.setNull(column);
		}

		@Override
		public void visitRowNumber(Column column, int index1) {
			pageBuilder.setBoolean(column, index1 != 0);
		}

		@Override
		public void visitColumnNumber(Column column, int index1) {
			pageBuilder.setBoolean(column, index1 != 0);
		}
	};

	protected class LongCellVisitor extends CellVisitor {

		@Override
		public void visitCellValueNumeric(Column column, Object cell, double value) {
			pageBuilder.setLong(column, (long) value);
		}

		@Override
		public void visitCellValueString(Column column, Object cell, String value) {
			pageBuilder.setLong(column, Long.parseLong(value));
		}

		@Override
		public void visitCellValueBoolean(Column column, Object cell, boolean value) {
			pageBuilder.setLong(column, value ? 1 : 0);
		}

		@Override
		public void visitCellValueError(Column column, Object cell, int code) {
			pageBuilder.setLong(column, code);
		}

		@Override
		public void visitRowNumber(Column column, int index1) {
			pageBuilder.setLong(column, index1);
		}

		@Override
		public void visitColumnNumber(Column column, int index1) {
			pageBuilder.setLong(column, index1);
		}
	};

	protected class DoubleCellVisitor extends CellVisitor {

		@Override
		public void visitCellValueNumeric(Column column, Object cell, double value) {
			pageBuilder.setDouble(column, value);
		}

		@Override
		public void visitCellValueString(Column column, Object cell, String value) {
			pageBuilder.setDouble(column, Double.parseDouble(value));
		}

		@Override
		public void visitCellValueBoolean(Column column, Object cell, boolean value) {
			pageBuilder.setDouble(column, value ? 1 : 0);
		}

		@Override
		public void visitCellValueError(Column column, Object cell, int code) {
			pageBuilder.setDouble(column, code);
		}

		@Override
		public void visitRowNumber(Column column, int index1) {
			pageBuilder.setDouble(column, index1);
		}

		@Override
		public void visitColumnNumber(Column column, int index1) {
			pageBuilder.setDouble(column, index1);
		}
	};

	protected class StringCellVisitor extends CellVisitor {

		@Override
		public void visitCellValueNumeric(Column column, Object cell, double value) {
			String s = Double.toString(value);
			if (s.endsWith(".0")) {
				s = s.substring(0, s.length() - 2);
			}
			pageBuilder.setString(column, s);
		}

		@Override
		public void visitCellValueString(Column column, Object cell, String value) {
			pageBuilder.setString(column, value);
		}

		@Override
		public void visitCellValueBoolean(Column column, Object cell, boolean value) {
			pageBuilder.setString(column, Boolean.toString(value));
		}

		@Override
		public void visitCellValueError(Column column, Object cell, int code) {
			FormulaError error = FormulaError.forInt((byte) code);
			String value = error.getString();
			pageBuilder.setString(column, value);
		}

		@Override
		public void visitRowNumber(Column column, int index1) {
			pageBuilder.setString(column, Integer.toString(index1));
		}

		@Override
		public void visitColumnNumber(Column column, int index1) {
			String value = CellReference.convertNumToColString(index1 - 1);
			pageBuilder.setString(column, value);
		}
	};

	protected class TimestampCellVisitor extends CellVisitor {

		@Override
		public void visitCellValueNumeric(Column column, Object cell, double value) {
			TimestampParser parser = getTimestampParser(column);
			TimeZone tz = parser.getDefaultTimeZone().toTimeZone();
			Date date = DateUtil.getJavaDate(value, tz);
			pageBuilder.setTimestamp(column, Timestamp.ofEpochMilli(date.getTime()));
		}

		@Override
		public void visitCellValueString(Column column, Object cell, String value) {
			TimestampParser parser = getTimestampParser(column);
			pageBuilder.setTimestamp(column, parser.parse(value));
		}

		@Override
		public void visitCellValueBoolean(Column column, Object cell, boolean value) {
			throw new UnsupportedOperationException("unsupported conversion Excel boolean to Embulk timestamp.");
		}

		@Override
		public void visitCellValueError(Column column, Object cell, int code) {
			pageBuilder.setNull(column);
		}

		@Override
		public void visitRowNumber(Column column, int index1) {
			throw new UnsupportedOperationException("unsupported conversion row_number to Embulk timestamp.");
		}

		@Override
		public void visitColumnNumber(Column column, int index1) {
			throw new UnsupportedOperationException("unsupported conversion column_number to Embulk timestamp.");
		}
	};

	protected final void visitCell0(Column column, CellVisitor visitor) {
		try {
			visitCell(column, visitor);
		} catch (Exception e) {
			throw new RuntimeException(MessageFormat.format("error {0} cell={1}!{2}. {3}", column,
					sheet.getSheetName(), new CellReference(currentRow.getRowNum(), getColumnOption(column)
							.getColumnIndex()).formatAsString(), e.getMessage()), e);
		}
	}

	protected void visitCell(Column column, CellVisitor visitor) {
		ColumnOptionTask option = getColumnOption(column);
		PoiExcelColumnValueType valueType = option.getValueTypeEnum();

		switch (valueType) {
		case SHEET_NAME:
			visitor.visitSheetName(column, sheet.getSheetName());
			return;
		case ROW_NUMBER:
			visitor.visitRowNumber(column, currentRow.getRowNum() + 1);
			return;
		case COLUMN_NUMBER:
			visitor.visitColumnNumber(column, option.getColumnIndex() + 1);
			return;
		default:
			break;
		}

		assert valueType.useCell();
		Cell cell = currentRow.getCell(option.getColumnIndex());
		if (cell == null) {
			visitCellNull(column);
			return;
		}
		switch (valueType) {
		case CELL_VALUE:
		case CELL_FORMULA:
			visitCellValue(column, option, cell, visitor);
			return;
		case CELL_STYLE:
			visitCellStyle(column, option, cell, visitor);
			return;
		case CELL_COMMENT:
			visitCellComment(column, option, cell, visitor);
			return;
		default:
			throw new UnsupportedOperationException(MessageFormat.format("unsupported value_type={0}", valueType));
		}
	}

	protected void visitCellNull(Column column) {
		pageBuilder.setNull(column);
	}

	protected void visitCellValue(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
		assert cell != null;

		int cellType = cell.getCellType();
		switch (cellType) {
		case Cell.CELL_TYPE_NUMERIC:
			visitor.visitCellValueNumeric(column, cell, cell.getNumericCellValue());
			return;
		case Cell.CELL_TYPE_STRING:
			visitor.visitCellValueString(column, cell, cell.getStringCellValue());
			return;
		case Cell.CELL_TYPE_FORMULA:
			PoiExcelColumnValueType valueType = option.getValueTypeEnum();
			if (valueType == PoiExcelColumnValueType.CELL_FORMULA) {
				visitor.visitCellFormula(column, cell);
			} else {
				visitCellValueFormula(column, option, cell, visitor);
			}
			return;
		case Cell.CELL_TYPE_BLANK:
			visitCellValueBlank(column, option, cell, visitor);
			return;
		case Cell.CELL_TYPE_BOOLEAN:
			visitor.visitCellValueBoolean(column, cell, cell.getBooleanCellValue());
			return;
		case Cell.CELL_TYPE_ERROR:
			visitCellValueError(column, option, cell, cell.getErrorCellValue(), visitor);
			return;
		default:
			throw new IllegalStateException(MessageFormat.format("unsupported POI cellType={0}", cellType));
		}
	}

	protected void visitCellValueBlank(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
		assert cell.getCellType() == Cell.CELL_TYPE_BLANK;

		boolean search = option.getSearchMergedCell().or(task.getSearchMergedCell());
		if (!search) {
			visitor.visitCellValueBlank(column, cell);
			return;
		}

		int r = cell.getRowIndex();
		int c = cell.getColumnIndex();

		Sheet s = cell.getSheet();
		int size = s.getNumMergedRegions();
		for (int i = 0; i < size; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			if (range.isInRange(r, c)) {
				Row firstRow = s.getRow(range.getFirstRow());
				if (firstRow == null) {
					visitCellNull(column);
					return;
				}
				Cell firstCell = firstRow.getCell(range.getFirstColumn());
				if (firstCell == null) {
					visitCellNull(column);
					return;
				}

				visitCellValue(column, option, firstCell, visitor);
				return;
			}
		}

		visitor.visitCellValueBlank(column, cell);
	}

	protected void visitCellValueFormula(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
		assert cell.getCellType() == Cell.CELL_TYPE_FORMULA;

		Optional<List<FormulaReplaceTask>> replaceOption = option.getFormulaReplace();
		if (!replaceOption.isPresent()) {
			replaceOption = task.getFormulaReplace();
		}
		if (replaceOption.isPresent()) {
			String formula = cell.getCellFormula();
			String old = formula;

			List<FormulaReplaceTask> list = replaceOption.get();
			for (FormulaReplaceTask replace : list) {
				String regex = replace.getRegex();
				String replacement = replace.getTo();

				replacement = replacement.replace("${row}", Integer.toString(cell.getRowIndex() + 1));

				formula = formula.replaceAll(regex, replacement);
			}

			if (!formula.equals(old)) {
				log.debug("formula replaced. old=\"{}\", new=\"{}\"", old, formula);
				try {
					cell.setCellFormula(formula);
				} catch (Exception e) {
					throw new RuntimeException(MessageFormat.format("setCellFormula error. formula={0}", formula), e);
				}
			}
		}

		CellValue cellValue;
		try {
			Workbook book = cell.getSheet().getWorkbook();
			CreationHelper helper = book.getCreationHelper();
			FormulaEvaluator evaluator = helper.createFormulaEvaluator();
			cellValue = evaluator.evaluate(cell);
		} catch (Exception e) {
			boolean setNull = option.getFormulaErrorNull().or(task.getFormulaErrorNull());
			if (setNull) {
				pageBuilder.setNull(column);
				return;
			}

			throw new RuntimeException(MessageFormat.format("evaluate error. formula={0}", cell.getCellFormula()), e);
		}

		int cellType = cellValue.getCellType();
		switch (cellType) {
		case Cell.CELL_TYPE_NUMERIC:
			visitor.visitCellValueNumeric(column, cellValue, cellValue.getNumberValue());
			return;
		case Cell.CELL_TYPE_STRING:
			visitor.visitCellValueString(column, cellValue, cellValue.getStringValue());
			return;
		case Cell.CELL_TYPE_BLANK:
			visitor.visitCellValueBlank(column, cellValue);
			return;
		case Cell.CELL_TYPE_BOOLEAN:
			visitor.visitCellValueBoolean(column, cellValue, cellValue.getBooleanValue());
			return;
		case Cell.CELL_TYPE_ERROR:
			visitCellValueError(column, option, cellValue, cellValue.getErrorValue(), visitor);
			return;
		case Cell.CELL_TYPE_FORMULA:
		default:
			throw new IllegalStateException(MessageFormat.format("unsupported POI cellType={0}", cellType));
		}
	}

	protected void visitCellValueError(Column column, ColumnOptionTask option, Object cell, int errorCode,
			CellVisitor visitor) {
		boolean setNull = option.getCellErrorNull().or(task.getCellErrorNull());
		if (setNull) {
			pageBuilder.setNull(column);
			return;
		}

		visitor.visitCellValueError(column, cell, errorCode);
	}

	protected void visitCellStyle(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
		Optional<String> nameOption = option.getCellStyleName();
		if (!nameOption.isPresent()) {
			throw new RuntimeException(MessageFormat.format("cell_style_name must be specified. column.name={0}",
					column.getName()));
		}

		String name = nameOption.get();
		CellStyle style = cell.getCellStyle();
		if (name.startsWith("font")) {
			visitCellStyleFont(column, style, name);
			return;
		}
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

	protected void visitCellStyleFont(Column column, CellStyle style, String name) {
		short index = style.getFontIndex();
		if (name.equals("font") && column.getType() instanceof LongType) {
			pageBuilder.setLong(column, index);
			return;
		}
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

	protected void visitCellColor(Column column, short colorIndex) {
		HSSFWorkbook book = (HSSFWorkbook) sheet.getWorkbook();
		HSSFPalette palette = book.getCustomPalette();
		HSSFColor color = palette.getColor(colorIndex);
		visitCellColor(column, color);
	}

	protected void visitCellColor(Column column, Color color) {
		if (color == null) {
			pageBuilder.setNull(column);
			return;
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

		if (column.getType() instanceof StringType) {
			String s = String.format("%02x%02x%02x", rgb[0], rgb[1], rgb[2]);
			pageBuilder.setString(column, s);
		} else {
			long n = (rgb[0] << 16) | (rgb[1] << 8) | rgb[2];
			pageBuilder.setLong(column, n);
		}
	}

	protected void visitCellComment(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
		Comment comment = cell.getCellComment();
		if (comment == null) {
			pageBuilder.setNull(column);
			return;
		}
		String s = comment.getString().getString();
		pageBuilder.setString(column, s);
	}
}
