package com.aires.excel.resolver;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.aires.excel.util.ExcelUtil;

/**
 * @author aires
 * 2017年11月20日 下午5:55:40
 * 描述：简单实现
 */
public abstract class SimpleRowResolver implements RowResolver{

	@Override
	public void resolve(int rowNum, Row row) {
		Object[]values = new Object[row.getLastCellNum() - row.getFirstCellNum()];
		int index = 0;
		Iterator<Cell>it = row.cellIterator();
		boolean hasNotEmptyValue = false, skipRow = skipEmptyRow();
		while(it.hasNext()){
			values[index] = ExcelUtil.val(it.next());
			if(skipRow && !hasNotEmptyValue && values[index] != null && String.valueOf(values[index]).trim().length() > 0)hasNotEmptyValue = true;
			
			index++;
		}
		if(skipRow && !hasNotEmptyValue)return;
		
		resolve(rowNum, values);
	}
	/**
	 * 
	 * @param rowNum 行索引
	 * @param values 值参考See Also
	 * @see ExcelUtil#val(Cell cell)
	 */
	protected abstract void resolve(int rowNum, Object[]values);

	/**
	 * 是否跳过空行
	 * @return
	 */
	protected boolean skipEmptyRow(){
		return true;
	}
}
