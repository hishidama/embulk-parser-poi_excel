package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.embulk.parser.poi_excel.PoiExcelColumnValueType;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.FormulaReplaceTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.Exec;
import org.embulk.spi.PageBuilder;
import org.slf4j.Logger;

import com.google.common.base.Optional;

public class PoiExcelCellVisitor {
	private final Logger log = Exec.getLogger(getClass());

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public PoiExcelCellVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visitCellValue(Column column, ColumnOptionTask option, Cell cell, CellVisitor visitor) {
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

		PluginTask task = visitorValue.getPluginTask();
		boolean search = option.getSearchMergedCell().or(task.getSearchMergedCell());
		if (!search) {
			visitor.visitCellValueBlank(column, cell);
			return;
		}

		int r = cell.getRowIndex();
		int c = cell.getColumnIndex();

		Sheet sheet = cell.getSheet();
		int size = sheet.getNumMergedRegions();
		for (int i = 0; i < size; i++) {
			CellRangeAddress range = visitorValue.getSheet().getMergedRegion(i);
			if (range.isInRange(r, c)) {
				Row firstRow = sheet.getRow(range.getFirstRow());
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
			PluginTask task = visitorValue.getPluginTask();
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
			PluginTask task = visitorValue.getPluginTask();
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
		PluginTask task = visitorValue.getPluginTask();
		boolean setNull = option.getCellErrorNull().or(task.getCellErrorNull());
		if (setNull) {
			pageBuilder.setNull(column);
			return;
		}

		visitor.visitCellValueError(column, cell, errorCode);
	}

	protected void visitCellNull(Column column) {
		pageBuilder.setNull(column);
	}
}
