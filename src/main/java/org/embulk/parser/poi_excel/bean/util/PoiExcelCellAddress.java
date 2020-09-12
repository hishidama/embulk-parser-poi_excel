package org.embulk.parser.poi_excel.bean.util;

import java.text.MessageFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

public class PoiExcelCellAddress {
	private final CellReference cellReference;

	public PoiExcelCellAddress(String cellAddress) {
		this.cellReference = new CellReference(cellAddress);
	}

	public String getSheetName() {
		return cellReference.getSheetName();
	}

	public Sheet getSheet(Row currentRow) {
		String sheetName = getSheetName();
		if (sheetName != null) {
			Workbook book = currentRow.getSheet().getWorkbook();
			Sheet sheet = book.getSheet(sheetName);
			if (sheet == null) {
				throw new RuntimeException(MessageFormat.format("not found sheet. sheetName={0}", sheetName));
			}
			return sheet;
		} else {
			return currentRow.getSheet();
		}
	}

	public Cell getCell(Row currentRow) {
		Sheet sheet = getSheet(currentRow);

		Row row = sheet.getRow(cellReference.getRow());
		if (row == null) {
			return null;
		}

		return row.getCell(cellReference.getCol());
	}

	public String getString() {
		return cellReference.formatAsString();
	}
}
