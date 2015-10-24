package org.embulk.parser.poi_excel.visitor.embulk;

import java.util.Date;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorValue;
import org.embulk.spi.Column;
import org.embulk.spi.time.Timestamp;
import org.embulk.spi.time.TimestampParseException;
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
		Timestamp t = Timestamp.ofEpochMilli(date.getTime());
		pageBuilder.setTimestamp(column, t);
	}

	@Override
	public void visitCellValueString(Column column, Object source, String value) {
		Timestamp t;
		try {
			TimestampParser parser = getTimestampParser(column);
			t = parser.parse(value);
		} catch (TimestampParseException e) {
			doConvertError(column, value, e);
			return;
		}
		pageBuilder.setTimestamp(column, t);
	}

	@Override
	public void visitCellValueBoolean(Column column, Object source, boolean value) {
		doConvertError(column, value, new UnsupportedOperationException(
				"unsupported conversion Excel boolean to Embulk timestamp"));
	}

	@Override
	public void visitCellValueError(Column column, Object source, int code) {
		doConvertError(column, code, new UnsupportedOperationException(
				"unsupported conversion Excel Cell error code to Embulk timestamp"));
	}

	@Override
	public void visitValueLong(Column column, Object source, long value) {
		pageBuilder.setTimestamp(column, Timestamp.ofEpochMilli(value));
	}

	@Override
	public void visitSheetName(Column column) {
		Sheet sheet = visitorValue.getSheet();
		doConvertError(column, sheet.getSheetName(), new UnsupportedOperationException(
				"unsupported conversion sheet_name to Embulk timestamp"));
	}

	@Override
	public void visitRowNumber(Column column, int index1) {
		doConvertError(column, index1, new UnsupportedOperationException(
				"unsupported conversion row_number to Embulk timestamp"));
	}

	@Override
	public void visitColumnNumber(Column column, int index1) {
		doConvertError(column, index1, new UnsupportedOperationException(
				"unsupported conversion column_number to Embulk timestamp"));
	}

	@Override
	protected void doConvertErrorConstant(Column column, String value) throws Exception {
		TimestampParser parser = getTimestampParser(column);
		pageBuilder.setTimestamp(column, parser.parse(value));
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
