package org.embulk.parser.poi_excel.bean.util;

import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.util.CellRangeAddress;
import org.embulk.parser.poi_excel.visitor.util.MergedRegionFinder;
import org.embulk.parser.poi_excel.visitor.util.MergedRegionList;
import org.embulk.parser.poi_excel.visitor.util.MergedRegionMap;
import org.embulk.parser.poi_excel.visitor.util.MergedRegionNothing;

public enum SearchMergedCell {
	NONE {
		@Override
		public MergedRegionFinder createMergedRegionFinder() {
			return new MergedRegionNothing();
		}
	},
	LINEAR_SEARCH {
		@Override
		public MergedRegionFinder createMergedRegionFinder() {
			return new MergedRegionList();
		}
	},
	TREE_SEARCH {
		@Override
		public MergedRegionFinder createMergedRegionFinder() {
			return new MergedRegionMap() {

				@Override
				protected Map<Integer, Map<Integer, CellRangeAddress>> newRowMap() {
					return new TreeMap<>();
				}

				@Override
				protected Map<Integer, CellRangeAddress> newColumnMap() {
					return new TreeMap<>();
				}
			};
		}
	},
	HASH_SEARCH {
		@Override
		public MergedRegionFinder createMergedRegionFinder() {
			return new MergedRegionMap() {

				@Override
				protected Map<Integer, Map<Integer, CellRangeAddress>> newRowMap() {
					return new HashMap<>();
				}

				@Override
				protected Map<Integer, CellRangeAddress> newColumnMap() {
					return new HashMap<>();
				}
			};
		}
	};

	private MergedRegionFinder mergedRegionFinder;

	public MergedRegionFinder getMergedRegionFinder() {
		if (mergedRegionFinder == null) {
			this.mergedRegionFinder = createMergedRegionFinder();
		}
		return mergedRegionFinder;
	}

	protected abstract MergedRegionFinder createMergedRegionFinder();
}