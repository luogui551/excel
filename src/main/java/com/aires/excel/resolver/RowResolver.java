/**
 * 
 */
package com.aires.excel.resolver;

import org.apache.poi.ss.usermodel.Row;

/**
 * 
 * @author aires
 * 2016-9-20 下午5:57:08
 * 描述：行解析器
 */
public interface RowResolver {
	/**
	 * 处理一行
	 * @param rowNum
	 * @param row
	 */
	public void resolve(int rowNum, Row row);
}
