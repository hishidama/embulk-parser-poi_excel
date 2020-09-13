package org.embulk.parser.poi_excel.visitor.util;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class MergedRegionNothing implements MergedRegionFinder {

	@Override
	public CellRangeAddress get(Sheet sheet, int rowIndex, int columnIndex) {
		return null;
	}
}
