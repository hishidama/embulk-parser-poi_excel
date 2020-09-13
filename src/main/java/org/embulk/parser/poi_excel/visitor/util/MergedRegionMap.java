package org.embulk.parser.poi_excel.visitor.util;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public abstract class MergedRegionMap implements MergedRegionFinder {

	private final Map<Sheet, Map<Integer, Map<Integer, CellRangeAddress>>> sheetMap = new ConcurrentHashMap<>();

	@Override
	public CellRangeAddress get(Sheet sheet, int rowIndex, int columnIndex) {
		Map<Integer, Map<Integer, CellRangeAddress>> rowMap = sheetMap.get(sheet);
		if (rowMap == null) {
			synchronized (sheet) {
				rowMap = createRowMap(sheet);
				sheetMap.put(sheet, rowMap);
			}
		}

		Map<Integer, CellRangeAddress> columnMap = rowMap.get(rowIndex);
		if (columnMap == null) {
			return null;
		}
		return columnMap.get(columnIndex);
	}

	protected Map<Integer, Map<Integer, CellRangeAddress>> createRowMap(Sheet sheet) {
		Map<Integer, Map<Integer, CellRangeAddress>> rowMap = newRowMap();

		for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
			CellRangeAddress region = sheet.getMergedRegion(i);

			for (int r = region.getFirstRow(); r <= region.getLastRow(); r++) {
				Map<Integer, CellRangeAddress> columnMap = rowMap.get(r);
				if (columnMap == null) {
					columnMap = newColumnMap();
					rowMap.put(r, columnMap);
				}

				for (int c = region.getFirstColumn(); c <= region.getLastColumn(); c++) {
					columnMap.put(c, region);
				}
			}
		}

		return rowMap;
	}

	protected abstract Map<Integer, Map<Integer, CellRangeAddress>> newRowMap();

	protected abstract Map<Integer, CellRangeAddress> newColumnMap();
}
