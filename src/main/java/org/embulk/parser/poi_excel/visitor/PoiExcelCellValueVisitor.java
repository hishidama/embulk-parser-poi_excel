package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.embulk.parser.poi_excel.PoiExcelColumnValueType;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.FormulaReplaceTask;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean.ErrorStrategy;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean.FormulaHandling;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.parser.poi_excel.visitor.util.MergedRegionFinder;
import org.embulk.spi.Column;
import org.embulk.spi.Exec;
import org.embulk.spi.PageBuilder;
import org.slf4j.Logger;

public class PoiExcelCellValueVisitor {
	private final Logger log = Exec.getLogger(getClass());

	protected final PoiExcelVisitorValue visitorValue;
	protected final PageBuilder pageBuilder;

	public PoiExcelCellValueVisitor(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		this.pageBuilder = visitorValue.getPageBuilder();
	}

	public void visitCellValue(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
		assert cell != null;

		Column column = bean.getColumn();

		CellType cellType = cell.getCellType();
		switch (cellType) {
		case NUMERIC:
			visitor.visitCellValueNumeric(column, cell, cell.getNumericCellValue());
			return;
		case STRING:
			visitor.visitCellValueString(column, cell, cell.getStringCellValue());
			return;
		case FORMULA:
			PoiExcelColumnValueType valueType = bean.getValueType();
			if (valueType == PoiExcelColumnValueType.CELL_FORMULA) {
				visitor.visitCellFormula(column, cell);
			} else {
				visitCellValueFormula(bean, cell, visitor);
			}
			return;
		case BLANK:
			visitCellValueBlank(bean, cell, visitor);
			return;
		case BOOLEAN:
			visitor.visitCellValueBoolean(column, cell, cell.getBooleanCellValue());
			return;
		case ERROR:
			visitCellValueError(bean, cell, cell.getErrorCellValue(), visitor);
			return;
		default:
			throw new IllegalStateException(MessageFormat.format("unsupported POI cellType={0}", cellType));
		}
	}

	protected void visitCellValueBlank(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
		assert cell.getCellType() == CellType.BLANK;

		Column column = bean.getColumn();

		CellRangeAddress region = findRegion(bean, cell);
		if (region != null) {
			Row firstRow = cell.getSheet().getRow(region.getFirstRow());
			if (firstRow == null) {
				visitCellNull(column);
				return;
			}
			Cell firstCell = firstRow.getCell(region.getFirstColumn());
			if (firstCell == null) {
				visitCellNull(column);
				return;
			}

			if (firstCell.getRowIndex() != cell.getRowIndex() || firstCell.getColumnIndex() != cell.getColumnIndex()) {
				visitCellValue(bean, firstCell, visitor);
				return;
			}
		}

		visitor.visitCellValueBlank(column, cell);
	}

	protected CellRangeAddress findRegion(PoiExcelColumnBean bean, Cell cell) {
		Sheet sheet = cell.getSheet();
		int r = cell.getRowIndex();
		int c = cell.getColumnIndex();

		MergedRegionFinder finder = bean.getMergedRegionFinder();
		return finder.get(sheet, r, c);
	}

	protected void visitCellValueFormula(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
		assert cell.getCellType() == CellType.FORMULA;

		FormulaHandling handling = bean.getFormulaHandling();
		switch (handling) {
		case CASHED_VALUE:
			visitCellValueFormulaCashedValue(bean, cell, visitor);
			break;
		default:
			visitCellValueFormulaEvaluate(bean, cell, visitor);
			break;
		}
	}

	protected void visitCellValueFormulaCashedValue(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
		Column column = bean.getColumn();

		CellType cellType = cell.getCachedFormulaResultType();
		switch (cellType) {
		case NUMERIC:
			visitor.visitCellValueNumeric(column, cell, cell.getNumericCellValue());
			return;
		case STRING:
			visitor.visitCellValueString(column, cell, cell.getStringCellValue());
			return;
		case BLANK:
			visitCellValueBlank(bean, cell, visitor);
			return;
		case BOOLEAN:
			visitor.visitCellValueBoolean(column, cell, cell.getBooleanCellValue());
			return;
		case ERROR:
			visitCellValueError(bean, cell, cell.getErrorCellValue(), visitor);
			return;
		case FORMULA:
		default:
			throw new IllegalStateException(MessageFormat.format("unsupported POI cellType={0}", cellType));
		}
	}

	protected void visitCellValueFormulaEvaluate(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
		Column column = bean.getColumn();

		List<FormulaReplaceTask> list = bean.getFormulaReplace();
		if (!list.isEmpty()) {
			String formula = cell.getCellFormula();
			String old = formula;

			for (FormulaReplaceTask replace : list) {
				String regex = replace.getRegex();
				String replacement = replace.getTo();

				if (replacement.contains("${row}")) {
					replacement = replacement.replace("${row}", Integer.toString(cell.getRowIndex() + 1));
				}
				if (replacement.contains("${column}")) {
					replacement = replacement.replace("${column}",
							CellReference.convertNumToColString(cell.getColumnIndex() + 1));
				}

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
			ErrorStrategy strategy = bean.getEvaluateErrorStrategy();
			switch (strategy.getStrategy()) {
			default:
				break;
			case CONSTANT:
				String value = strategy.getValue();
				if (value == null) {
					pageBuilder.setNull(column);
				} else {
					visitor.visitCellValueString(column, cell, value);
				}
				return;
			}

			throw new RuntimeException(MessageFormat.format("evaluate error. formula={0}", cell.getCellFormula()), e);
		}

		CellType cellType = cellValue.getCellType();
		switch (cellType) {
		case NUMERIC:
			visitor.visitCellValueNumeric(column, cellValue, cellValue.getNumberValue());
			return;
		case STRING:
			visitor.visitCellValueString(column, cellValue, cellValue.getStringValue());
			return;
		case BLANK:
			visitor.visitCellValueBlank(column, cellValue);
			return;
		case BOOLEAN:
			visitor.visitCellValueBoolean(column, cellValue, cellValue.getBooleanValue());
			return;
		case ERROR:
			visitCellValueError(bean, cellValue, cellValue.getErrorValue(), visitor);
			return;
		case FORMULA:
		default:
			throw new IllegalStateException(MessageFormat.format("unsupported POI cellType={0}", cellType));
		}
	}

	protected void visitCellValueError(PoiExcelColumnBean bean, Object cell, int errorCode, CellVisitor visitor) {
		Column column = bean.getColumn();

		ErrorStrategy strategy = bean.getCellErrorStrategy();
		switch (strategy.getStrategy()) {
		default:
			pageBuilder.setNull(column);
			return;
		case CONSTANT:
			String value = strategy.getValue();
			if (value == null) {
				pageBuilder.setNull(column);
			} else {
				visitor.visitCellValueString(column, cell, value);
			}
			return;
		case ERROR_CODE:
			break;
		case EXCEPTION:
			FormulaError error = FormulaError.forInt((byte) errorCode);
			throw new RuntimeException(MessageFormat.format("encount cell error. error_code={0}({1})", errorCode,
					error.getString()));
		}

		visitor.visitCellValueError(column, cell, errorCode);
	}

	protected void visitCellNull(Column column) {
		pageBuilder.setNull(column);
	}
}
