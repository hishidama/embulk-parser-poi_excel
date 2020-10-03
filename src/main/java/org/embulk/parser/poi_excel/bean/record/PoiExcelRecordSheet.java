package org.embulk.parser.poi_excel.bean.record;

import org.apache.poi.ss.usermodel.Cell;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.spi.Exec;
import org.slf4j.Logger;

public class PoiExcelRecordSheet extends PoiExcelRecord {
	private final Logger log = Exec.getLogger(getClass());

	private boolean exists;

	@Override
	protected void initializeLoop(int skipHeaderLines) {
		this.exists = true;
	}

	@Override
	public boolean exists() {
		return exists;
	}

	@Override
	public void moveNext() {
		this.exists = false;
	}

	@Override
	protected void logStartEnd(String part) {
		if (log.isDebugEnabled()) {
			log.debug("sheet({}) {}", getSheet().getSheetName(), part);
		}
	}

	@Override
	public int getRowIndex(PoiExcelColumnBean bean) {
		throw new UnsupportedOperationException("unsupported at record_type=sheet");
	}

	@Override
	public int getColumnIndex(PoiExcelColumnBean bean) {
		throw new UnsupportedOperationException("unsupported at record_type=sheet");
	}

	@Override
	public Cell getCell(PoiExcelColumnBean bean) {
		throw new UnsupportedOperationException("unsupported at record_type=sheet");
	}
}
