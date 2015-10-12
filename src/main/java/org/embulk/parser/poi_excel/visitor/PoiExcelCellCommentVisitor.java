package org.embulk.parser.poi_excel.visitor;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.ColumnOptionTask;
import org.embulk.spi.Column;

public class PoiExcelCellCommentVisitor extends AbstractPoiExcelCellAttributeVisitor<Comment> {

	public PoiExcelCellCommentVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	protected Comment getAttributeSource(Column column, ColumnOptionTask option, Cell cell) {
		return cell.getCellComment();
	}

	@Override
	protected Map<String, AttributeSupplier<Comment>> getAttributeSupplierMap() {
		return SUPPLIER_MAP;
	}

	protected static final Map<String, AttributeSupplier<Comment>> SUPPLIER_MAP;
	static {
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

		// TODO getClientAnchor

		SUPPLIER_MAP = Collections.unmodifiableMap(map);
	}
}
