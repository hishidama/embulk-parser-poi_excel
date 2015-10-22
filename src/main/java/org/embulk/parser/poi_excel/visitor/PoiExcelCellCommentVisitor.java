package org.embulk.parser.poi_excel.visitor;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.spi.Column;

public class PoiExcelCellCommentVisitor extends AbstractPoiExcelCellAttributeVisitor<Comment> {

	public PoiExcelCellCommentVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	protected Comment getAttributeSource(Column column, PoiExcelColumnBean bean, Cell cell) {
		return cell.getCellComment();
	}

	protected boolean acceptKey(String key) {
		if (key.equals("client_anchor")) {
			return false;
		}
		return true;
	}

	@Override
	protected Map<String, AttributeSupplier<Comment>> getAttributeSupplierMap() {
		return SUPPLIER_MAP;
	}

	private final Map<String, AttributeSupplier<Comment>> SUPPLIER_MAP;
	{
		Map<String, AttributeSupplier<Comment>> map = new HashMap<>(32);
		map.put("author", new AttributeSupplier<Comment>() {
			@Override
			public Object get(Column column, Cell cell, Comment comment) {
				return comment.getAuthor();
			}
		});
		map.put("column", new AttributeSupplier<Comment>() {
			@Override
			public Object get(Column column, Cell cell, Comment comment) {
				return (long) comment.getColumn();
			}
		});
		map.put("row", new AttributeSupplier<Comment>() {
			@Override
			public Object get(Column column, Cell cell, Comment comment) {
				return (long) comment.getRow();
			}
		});
		map.put("is_visible", new AttributeSupplier<Comment>() {
			@Override
			public Object get(Column column, Cell cell, Comment comment) {
				return comment.isVisible();
			}
		});
		map.put("string", new AttributeSupplier<Comment>() {
			@Override
			public Object get(Column column, Cell cell, Comment comment) {
				RichTextString rich = comment.getString();
				return rich.getString();
			}
		});
		map.put("client_anchor", new AttributeSupplier<Comment>() {
			@Override
			public Object get(Column column, Cell cell, Comment comment) {
				return getClientAnchorValue(column, cell, comment, null);
			}
		});
		for (String key : PoiExcelClientAnchorVisitor.getKeys()) {
			map.put("client_anchor." + key, new AttributeSupplier<Comment>() {
				@Override
				public Object get(Column column, Cell cell, Comment comment) {
					return getClientAnchorValue(column, cell, comment, key);
				}
			});
		}
		SUPPLIER_MAP = Collections.unmodifiableMap(map);
	}

	final Object getClientAnchorValue(Column column, Cell cell, Comment comment, String key) {
		ClientAnchor anchor = comment.getClientAnchor();
		PoiExcelVisitorFactory factory = visitorValue.getVisitorFactory();
		PoiExcelClientAnchorVisitor delegator = factory.getPoiExcelClientAnchorVisitor();
		return delegator.getClientAnchorValue(column, cell, anchor, key);
	}
}
