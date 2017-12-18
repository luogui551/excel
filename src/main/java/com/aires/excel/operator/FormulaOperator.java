package com.aires.excel.operator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

/**
 * @author aires
 * 2017年12月13日 下午2:53:15
 * 描述：公式单元格
 */
public class FormulaOperator extends CellOperator{

	@Override
	public Object get0() {
		System.out.println(CellOperator.getByType(getCell().getCachedFormulaResultTypeEnum()));
		return CellOperator.getByType(getCell().getCachedFormulaResultTypeEnum()).get();
	}

	@Override
	public void write(Object v) {
		Cell cell = getCell();
		cell.setCellType(CellType.FORMULA);
		cell.setCellValue((String)v);
	}

}
