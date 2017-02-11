package org.embulk.parser.poi_excel.visitor;

import org.embulk.parser.poi_excel.visitor.embulk.BooleanCellVisitor;
import org.embulk.parser.poi_excel.visitor.embulk.DoubleCellVisitor;
import org.embulk.parser.poi_excel.visitor.embulk.LongCellVisitor;
import org.embulk.parser.poi_excel.visitor.embulk.StringCellVisitor;
import org.embulk.parser.poi_excel.visitor.embulk.TimestampCellVisitor;

public class PoiExcelVisitorFactory {

	protected final PoiExcelVisitorValue visitorValue;

	public PoiExcelVisitorFactory(PoiExcelVisitorValue visitorValue) {
		this.visitorValue = visitorValue;
		visitorValue.setVisitorFactory(this);
	}

	public final PoiExcelVisitorValue getVisitorValue() {
		return visitorValue;
	}

	// visitor root (Embulk ColumnVisitor)
	private PoiExcelColumnVisitor poiExcelColumnVisitor;

	public final PoiExcelColumnVisitor getPoiExcelColumnVisitor() {
		if (poiExcelColumnVisitor == null) {
			poiExcelColumnVisitor = newPoiExcelColumnVisitor();
		}
		return poiExcelColumnVisitor;
	}

	protected PoiExcelColumnVisitor newPoiExcelColumnVisitor() {
		return new PoiExcelColumnVisitor(visitorValue);
	}

	// Embulk boolean
	private BooleanCellVisitor booleanCellVisitor;

	public final BooleanCellVisitor getBooleanCellVisitor() {
		if (booleanCellVisitor == null) {
			booleanCellVisitor = newBooleanCellVisitor();
		}
		return booleanCellVisitor;
	}

	protected BooleanCellVisitor newBooleanCellVisitor() {
		return new BooleanCellVisitor(visitorValue);
	}

	// Embulk long
	private LongCellVisitor longCellVisitor;

	public final LongCellVisitor getLongCellVisitor() {
		if (longCellVisitor == null) {
			longCellVisitor = newLongCellVisitor();
		}
		return longCellVisitor;
	}

	protected LongCellVisitor newLongCellVisitor() {
		return new LongCellVisitor(visitorValue);
	}

	// Embulk double
	private DoubleCellVisitor doubleCellVisitor;

	public final DoubleCellVisitor getDoubleCellVisitor() {
		if (doubleCellVisitor == null) {
			doubleCellVisitor = newDoubleCellVisitor();
		}
		return doubleCellVisitor;
	}

	protected DoubleCellVisitor newDoubleCellVisitor() {
		return new DoubleCellVisitor(visitorValue);
	}

	// Embulk string
	private StringCellVisitor stringCellVisitor;

	public final StringCellVisitor getStringCellVisitor() {
		if (stringCellVisitor == null) {
			stringCellVisitor = newStringCellVisitor();
		}
		return stringCellVisitor;
	}

	protected StringCellVisitor newStringCellVisitor() {
		return new StringCellVisitor(visitorValue);
	}

	// Embulk timestamp
	private TimestampCellVisitor timestampCellVisitor;

	public final TimestampCellVisitor getTimestampCellVisitor() {
		if (timestampCellVisitor == null) {
			timestampCellVisitor = newTimestampCellVisitor();
		}
		return timestampCellVisitor;
	}

	protected TimestampCellVisitor newTimestampCellVisitor() {
		return new TimestampCellVisitor(visitorValue);
	}

	// cell value/formula
	private PoiExcelCellValueVisitor poiExcelCellValueVisitor;

	public final PoiExcelCellValueVisitor getPoiExcelCellValueVisitor() {
		if (poiExcelCellValueVisitor == null) {
			poiExcelCellValueVisitor = newPoiExcelCellValueVisitor();
		}
		return poiExcelCellValueVisitor;
	}

	protected PoiExcelCellValueVisitor newPoiExcelCellValueVisitor() {
		return new PoiExcelCellValueVisitor(visitorValue);
	}

	// cell style
	private PoiExcelCellStyleVisitor poiExcelCellStyleVisitor;

	public final PoiExcelCellStyleVisitor getPoiExcelCellStyleVisitor() {
		if (poiExcelCellStyleVisitor == null) {
			poiExcelCellStyleVisitor = newPoiExcelCellStyleVisitor();
		}
		return poiExcelCellStyleVisitor;
	}

	protected PoiExcelCellStyleVisitor newPoiExcelCellStyleVisitor() {
		return new PoiExcelCellStyleVisitor(visitorValue);
	}

	// cell font
	private PoiExcelCellFontVisitor poiExcelCellFontVisitor;

	public final PoiExcelCellFontVisitor getPoiExcelCellFontVisitor() {
		if (poiExcelCellFontVisitor == null) {
			poiExcelCellFontVisitor = newPoiExcelCellFontVisitor();
		}
		return poiExcelCellFontVisitor;
	}

	protected PoiExcelCellFontVisitor newPoiExcelCellFontVisitor() {
		return new PoiExcelCellFontVisitor(visitorValue);
	}

	// cell comment
	private PoiExcelCellCommentVisitor poiExcelCellCommentVisitor;

	public final PoiExcelCellCommentVisitor getPoiExcelCellCommentVisitor() {
		if (poiExcelCellCommentVisitor == null) {
			poiExcelCellCommentVisitor = newPoiExcelCellCommentVisitor();
		}
		return poiExcelCellCommentVisitor;
	}

	protected PoiExcelCellCommentVisitor newPoiExcelCellCommentVisitor() {
		return new PoiExcelCellCommentVisitor(visitorValue);
	}

	// cell type
	private PoiExcelCellTypeVisitor poiExcelCellTypeVisitor;

	public final PoiExcelCellTypeVisitor getPoiExcelCellTypeVisitor() {
		if (poiExcelCellTypeVisitor == null) {
			poiExcelCellTypeVisitor = newPoiExcelCellTypeVisitor();
		}
		return poiExcelCellTypeVisitor;
	}

	protected PoiExcelCellTypeVisitor newPoiExcelCellTypeVisitor() {
		return new PoiExcelCellTypeVisitor(visitorValue);
	}

	// ClientAnchor
	private PoiExcelClientAnchorVisitor poiExcelClientAnchorVisitor;

	public final PoiExcelClientAnchorVisitor getPoiExcelClientAnchorVisitor() {
		if (poiExcelClientAnchorVisitor == null) {
			poiExcelClientAnchorVisitor = newPoiExcelClientAnchorVisitor();
		}
		return poiExcelClientAnchorVisitor;
	}

	protected PoiExcelClientAnchorVisitor newPoiExcelClientAnchorVisitor() {
		return new PoiExcelClientAnchorVisitor(visitorValue);
	}

	// color
	private PoiExcelColorVisitor poiExcelColorVisitor;

	public final PoiExcelColorVisitor getPoiExcelColorVisitor() {
		if (poiExcelColorVisitor == null) {
			poiExcelColorVisitor = newPoiExcelColorVisitor();
		}
		return poiExcelColorVisitor;
	}

	protected PoiExcelColorVisitor newPoiExcelColorVisitor() {
		return new PoiExcelColorVisitor(visitorValue);
	}
}
