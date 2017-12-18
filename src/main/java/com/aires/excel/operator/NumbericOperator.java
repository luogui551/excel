package com.aires.excel.operator;

import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

/**
 * @author aires
 * 2017年12月13日 下午3:03:59
 * 描述：数字类
 */
public class NumbericOperator extends CellOperator{

	/**
	 * @return 日期、时间返回Date对象，其它数字返回Double
	 */
	@Override
	protected Object get0() {
		Cell cell = getCell();
		short format = cell.getCellStyle().getDataFormat();
		//日期和时间处理为Date
		if(HSSFDateUtil.isCellDateFormatted(cell) 
				|| (format == 31 || format == 57 || format == 58)//日期
				|| (format == 20 || format == 32 || format == 183 )){//时间
			
			return cell.getDateCellValue();
		}
		return cell.getNumericCellValue();
	}

	@Override
	public void write(Object v) {
		Cell cell = getCell();
		cell.setCellType(CellType.NUMERIC);	
		if(v instanceof Date){
			CellStyle style = cell.getRow().getSheet().getWorkbook().createCellStyle();
			style.setDataFormat((short)31);
			cell.setCellStyle(style);

			cell.setCellValue((Date)v);
		}else{
			cell.setCellValue(((Number)v).doubleValue());
		}
	}
}
