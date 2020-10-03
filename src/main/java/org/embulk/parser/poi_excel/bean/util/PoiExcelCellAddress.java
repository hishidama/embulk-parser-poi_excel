package org.embulk.parser.poi_excel.bean.util;

import java.text.MessageFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.embulk.parser.poi_excel.bean.record.PoiExcelRecord;

public class PoiExcelCellAddress {
	private final CellReference cellReference;

	public PoiExcelCellAddress(CellReference cellReference) {
		this.cellReference = cellReference;
	}

	public String getSheetName() {
		return cellReference.getSheetName();
	}

	public Sheet getSheet(PoiExcelRecord record) {
		String sheetName = getSheetName();
		if (sheetName != null) {
			Workbook book = record.getSheet().getWorkbook();
			Sheet sheet = book.getSheet(sheetName);
			if (sheet == null) {
				throw new RuntimeException(MessageFormat.format("not found sheet. sheetName={0}", sheetName));
			}
			return sheet;
		} else {
			return record.getSheet();
		}
	}

	public int getRowIndex() {
		return cellReference.getRow();
	}

	public int getColumnIndex() {
		return cellReference.getCol();
	}

	public Cell getCell(PoiExcelRecord record) {
		Sheet sheet = getSheet(record);

		Row row = sheet.getRow(getRowIndex());
		if (row == null) {
			return null;
		}

		return row.getCell(getColumnIndex());
	}

	public String getString() {
		return cellReference.formatAsString();
	}
}
