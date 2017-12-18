package com.aires.excel.resolver;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author aires
 * 2017年12月13日 下午4:25:45
 * 描述： Cell适配 Row
 */
public class CellRowResolverAdapter implements RowResolver{

	private CellResolver resolver;
	public CellRowResolverAdapter(CellResolver resolver){
		this.resolver = resolver;
	}
	@Override
	public void resolve(int rowNum, Row row) {
		int colNum = 0;
		Iterator<Cell>it = row.cellIterator();
		while(it.hasNext()){
			resolver.resolve(rowNum, colNum++, it.next());
		}		
	}

}
