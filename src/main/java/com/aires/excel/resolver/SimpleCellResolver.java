package com.aires.excel.resolver;

import org.apache.poi.ss.usermodel.Cell;

import com.aires.excel.util.ExcelUtil;

/**
 * @author aires
 * 2017年11月20日 下午5:18:06
 * 描述：简单实现
 */
public abstract class SimpleCellResolver implements CellResolver{
	
	@Override
	public void resolve(int rowNum, int colNum, Cell cell) {
		resolve(rowNum, colNum, ExcelUtil.val(cell));	
	}
	/**
	 * 按默认规则获取单元格数据
	 * @param rowNum
	 * @param colNum
	 * @param value 值参考See Also
	 * @see ExcelUtil#val(Cell cell)
	 */
	protected abstract void resolve(int rowNum, int colNum, Object value);
}
