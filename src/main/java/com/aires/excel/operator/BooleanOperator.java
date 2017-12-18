package com.aires.excel.operator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

/**
 * @author aires
 * 2017年12月13日 下午3:13:16
 * 描述：
 */
public class BooleanOperator extends CellOperator{

	@Override
	protected Object get0() {
		return getCell().getBooleanCellValue();
	}

	@Override
	public void write(Object v) {
		Cell cell = getCell();
		cell.setCellType(CellType.BOOLEAN);
		cell.setCellValue((Boolean)v);
	}

}
