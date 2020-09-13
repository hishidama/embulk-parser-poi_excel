package org.embulk.parser.poi_excel.visitor.util;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class MergedRegionList implements MergedRegionFinder {

	@Override
	public CellRangeAddress get(Sheet sheet, int rowIndex, int columnIndex) {
		int size = sheet.getNumMergedRegions();
		for (int i = 0; i < size; i++) {
			CellRangeAddress region = sheet.getMergedRegion(i);
			if (region.isInRange(rowIndex, columnIndex)) {
				return region;
			}
		}

		return null;
	}
}
