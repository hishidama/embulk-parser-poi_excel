package org.embulk.parser.poi_excel.visitor;

import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.spi.Column;

public class PoiExcelClientAnchorVisitor extends AbstractPoiExcelCellAttributeVisitor<ClientAnchor> {

	public PoiExcelClientAnchorVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	public Object getClientAnchorValue(Column column, Cell cell, ClientAnchor anchor, String key) {
		if (key == null || key.isEmpty()) {
			return getAllValues(column, cell, anchor);
		}

		return getAttributeValue(column, cell, anchor, key);
	}

	@Override
	protected ClientAnchor getAttributeSource(PoiExcelColumnBean bean, Cell cell) {
		throw new UnsupportedOperationException();
	}

	@Override
	protected Map<String, AttributeSupplier<ClientAnchor>> getAttributeSupplierMap() {
		return SUPPLIER_MAP;
	}

	private static final Map<String, AttributeSupplier<ClientAnchor>> SUPPLIER_MAP;
	static {
		Map<String, AttributeSupplier<ClientAnchor>> map = new HashMap<>(16);
		map.put("anchor_type", new AttributeSupplier<ClientAnchor>() {
			@Override
			public Object get(Column column, Cell cell, ClientAnchor anchor) {
				return (long) anchor.getAnchorType().value;
			}
		});
		map.put("col1", new AttributeSupplier<ClientAnchor>() {
			@Override
			public Object get(Column column, Cell cell, ClientAnchor anchor) {
				return (long) anchor.getCol1();
			}
		});
		map.put("col2", new AttributeSupplier<ClientAnchor>() {
			@Override
			public Object get(Column column, Cell cell, ClientAnchor anchor) {
				return (long) anchor.getCol2();
			}
		});
		map.put("dx1", new AttributeSupplier<ClientAnchor>() {
			@Override
			public Object get(Column column, Cell cell, ClientAnchor anchor) {
				return (long) anchor.getDx1();
			}
		});
		map.put("dx2", new AttributeSupplier<ClientAnchor>() {
			@Override
			public Object get(Column column, Cell cell, ClientAnchor anchor) {
				return (long) anchor.getDx2();
			}
		});
		map.put("dy1", new AttributeSupplier<ClientAnchor>() {
			@Override
			public Object get(Column column, Cell cell, ClientAnchor anchor) {
				return (long) anchor.getDy1();
			}
		});
		map.put("dy2", new AttributeSupplier<ClientAnchor>() {
			@Override
			public Object get(Column column, Cell cell, ClientAnchor anchor) {
				return (long) anchor.getDy2();
			}
		});
		map.put("row1", new AttributeSupplier<ClientAnchor>() {
			@Override
			public Object get(Column column, Cell cell, ClientAnchor anchor) {
				return (long) anchor.getRow1();
			}
		});
		map.put("row2", new AttributeSupplier<ClientAnchor>() {
			@Override
			public Object get(Column column, Cell cell, ClientAnchor anchor) {
				return (long) anchor.getRow2();
			}
		});
		SUPPLIER_MAP = Collections.unmodifiableMap(map);
	}

	public static Collection<String> getKeys() {
		return SUPPLIER_MAP.keySet();
	}
}
