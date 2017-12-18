package com.aires.excel.operator;

import org.apache.poi.ss.usermodel.CellType;

/**
 * @author aires
 * 2017年12月13日 下午2:57:23
 * 描述：文本单元格
 */
public class StringOperator extends CellOperator {

	@Override
	protected Object get0() {
		return getCell().getStringCellValue();
	}

	@Override
	public void write(Object v) {
		getCell().setCellType(CellType.STRING);
		getCell().setCellValue(v.toString());
	}

}
