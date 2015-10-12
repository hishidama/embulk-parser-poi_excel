package org.embulk.parser.poi_excel.visitor.embulk;

import java.util.Date;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.DateUtil;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorValue;
import org.embulk.spi.Column;
import org.embulk.spi.time.Timestamp;
import org.embulk.spi.time.TimestampParser;
import org.embulk.spi.util.Timestamps;

public class TimestampCellVisitor extends CellVisitor {

	public TimestampCellVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	public void visitCellValueNumeric(Column column, Object source, double value) {
		TimestampParser parser = getTimestampParser(column);
		TimeZone tz = parser.getDefaultTimeZone().toTimeZone();
		Date date = DateUtil.getJavaDate(value, tz);
		pageBuilder.setTimestamp(column, Timestamp.ofEpochMilli(date.getTime()));
	}

	@Override
	public void visitCellValueString(Column column, Object source, String value) {
		TimestampParser parser = getTimestampParser(column);
		pageBuilder.setTimestamp(column, parser.parse(value));
	}

	@Override
	public void visitCellValueBoolean(Column column, Object source, boolean value) {
		throw new UnsupportedOperationException("unsupported conversion Excel boolean to Embulk timestamp.");
	}

	@Override
	public void visitCellValueError(Column column, Object source, int code) {
		pageBuilder.setNull(column);
	}

	@Override
	public void visitSheetName(Column column) {
		throw new UnsupportedOperationException("unsupported conversion sheet_name to Embulk timestamp.");
	}

	@Override
	public void visitRowNumber(Column column, int index1) {
		throw new UnsupportedOperationException("unsupported conversion row_number to Embulk timestamp.");
	}

	@Override
	public void visitColumnNumber(Column column, int index1) {
		throw new UnsupportedOperationException("unsupported conversion column_number to Embulk timestamp.");
	}

	private TimestampParser[] timestampParsers;

	protected final TimestampParser getTimestampParser(Column column) {
		if (timestampParsers == null) {
			PluginTask task = visitorValue.getPluginTask();
			timestampParsers = Timestamps.newTimestampColumnParsers(task, task.getColumns());
		}
		return timestampParsers[column.getIndex()];
	}
}
