/**
 * 
 */
package com.aires.excel.util;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.aires.excel.Excel;
import com.aires.excel.exception.ExcelException;
import com.aires.excel.operator.CellOperator;
import com.aires.util.CloseableUtil;

/**
 * @author aires
 * 2016-9-19 上午11:23:10
 * 描述：解析、生成excel工具类
 */
public class ExcelUtil {
	private static final ThreadLocal<Integer>sheetIndex = new ThreadLocal<Integer>();
	private static final ThreadLocal<Sheet>sheet = new ThreadLocal<Sheet>();
	/**
	 * 当前处理的sheet索引，从0开始(需要在Resolver中使用)
	 * @return -1表示还未开始处理
	 */
	public static int sheetIndex(){
		Integer r = sheetIndex.get();
		return r == null ? -1 : r;
	}
	/**
	 * 当前处理的sheet(需要在Resolver中使用)
	 * @return null表示还未开始处理
	 */
	public static Sheet sheet(){
		return sheet.get();
	}
	/**
	 * 简单的获取单元格数据
	 * @param cell
	 * @return 数字类单元格返回为Double, 日期/时间返回Date, 公式返回计算后的结果，其它返回String
	 */
	public static <T>T val(Cell cell){
		return CellOperator.getByCell(cell).get();
	}
	/**
	 * 构建excel对象
	 * @param excel
	 * @return
	 */
	public static Excel parse(String excel){
		return parse(new File(excel));
	}
	/**
	 * 构建excel对象
	 * @param excel
	 * @return
	 */
	public static Excel parse(File excel){
		return new Excel(excel);
	}
	/**
	 * 构建excel对象
	 * @param is
	 * @return
	 */
	public static Excel parse(InputStream is){
		return new Excel(is);
	}
	/**
	 * 创建新excel
	 * @param defaultSheetName
	 * @return
	 */
	public static Excel newExcel(String defaultSheetName){
		return new Excel(defaultSheetName);		
	}
	/**
	 * 构建对象(关闭is)
	 * @param is
	 * @return
	 */
	public static Workbook buildWorkBook(InputStream is){
		CachedInputStream cis = new CachedInputStream(is);
		try{
			try {
				return new HSSFWorkbook(cis);
			} catch (OfficeXmlFileException e){
				cis.reset();
				return new XSSFWorkbook(cis);
			}
		}catch(IOException e){
			throw new ExcelException(e);
		} finally {
			CloseableUtil.close(cis, is);
		}
	}
	//辅助工具
	private static class CachedInputStream extends InputStream{
		
		private static final int LEN = 512;
		
		private InputStream is;
		
		private int[]cache = new int[LEN];
		
		private int index;
		
		private boolean reset;
		
		public CachedInputStream(InputStream is){
			this.is = is;
		}

		@Override
		public int read() throws IOException {
			if(index < LEN){
				if(reset)return cache[index++];
				else {
					return cache[index++] = is.read();
				}					
			}
			index++;
			return is.read();
		}
		
		public synchronized void reset() throws IOException {
			index = 0;
			reset = true;
	    }

		@Override
		public void close() throws IOException {
			if(index != LEN || reset){
				cache = null;
				is.close();
			}
		}
	}
}
