/**
 * 
 */
package com.aires.excel.resolver;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 
 * @author aires
 * 2016-9-20 下午5:57:17
 * 描述：单元格解析器
 */
public interface CellResolver {
	/**
	 * 处理单个单元格
	 * @param rowNum
	 * @param colNum
	 * @param cell
	 */
	public void resolve(int rowNum, int colNum, Cell cell);
}