/**
 * 
 */
package com.aires.excel;

import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.aires.excel.exception.ExcelException;
import com.aires.excel.operator.CellOperator;
import com.aires.excel.resolver.CellResolver;
import com.aires.excel.resolver.CellRowResolverAdapter;
import com.aires.excel.resolver.RowResolver;
import com.aires.excel.util.ExcelUtil;
import com.aires.util.CloseableUtil;
/**
 * 
 * @author aires
 * 2016-9-21 下午4:26:39
 * 描述：用于封装简单的excel处理<br>
 * 所有行、列索引值从0开始
 */
public class Excel implements Closeable{
	//原始对象
	private Workbook wb;
	//当前sheet
	private Sheet sheet;
	//当前行
	private Row cur;
	//下一行、列索引
	int rowNum, colNum, sheetIndex, numOfSheets;
	//目标excel文件，用于save时
	private File excelFile;
	//是否设置过列宽度
//	private boolean widthSetted;
	
	/**
	 * 创建包含指定sheetName的对象
	 * @param sheetName
	 */
	public Excel(String sheetName){
		wb = new HSSFWorkbook();
		newSheet(sheetName);
	}
	/**
	 * 包装指定文件的excel
	 * @param excelFile
	 */
	public Excel(File excelFile) {
		this(getFileInputStream(excelFile));
		this.excelFile = excelFile;
	}
	/**
	 * 包装指定InputStream
	 * @param is
	 */
	public Excel(InputStream is){
		this(ExcelUtil.buildWorkBook(is));
	}
	/**
	 * 包装Workbook对象
	 * @param wb
	 */
	public Excel(Workbook wb){
		this.wb = wb;
		numOfSheets = wb.getNumberOfSheets();
		switchSheet(0);
	}
	/**
	 * 获取Workbook对象
	 * @return
	 */
	public Workbook getWorkbook(){
		return wb;
	}
	
	//---------------------部分-------------------------
	/**
	 * 新增sheet页
	 * @param sheetName
	 */
	public Excel newSheet(String sheetName){
		this.sheet = wb.createSheet(sheetName);
		rowNum = colNum = 0;
		cur = sheet.createRow(rowNum++);
		
		numOfSheets++;
		return this;
	}
	/**
	 * 设置列宽度
	 * @param startCol 开始的列索引(当widths.length == 0时，表示设置当前索引列的宽度)
	 * @param widths 设置startCol（含）列后widths.length列的宽度
	 * @return
	 */
	public Excel setWidths(int startCol, int...widths){
		if(widths.length == 0)
			sheet.setColumnWidth(colNum, calcWidth(startCol));
		else
			for(int i = 0, len = widths.length; i < len; i++)sheet.setColumnWidth(startCol + i, calcWidth(widths[i]));
		return this;
	}
	/**
	 * 值写到当前单元格(数组中一个元素一个单元格)
	 * @param values
	 */
	public Excel write(Object...values){
		for(Object v : values){
			Cell c = cur.createCell(colNum++);
			if(v != null)CellOperator.write(c, v);
		}		
		return this;
	}
	/**
	 * 值写到新行(另起一行)
	 * @param values
	 */
	public Excel writeRow(Object...values){
		if(colNum > 0)next();//空行就写到当前行
		write(values);
		return this;
	}
	/**
	 * 合并单元格
	 * @param startRow
	 * @param rowCount
	 * @param startCol
	 * @param colCount
	 * @return
	 */
	public Excel merge(int startRow, int rowCount, int startCol, int colCount){
		if(!(rowCount <= 1 && colCount <= 1)){
			sheet.addMergedRegion(new CellRangeAddress(startRow, calc(startRow, rowCount), startCol, calc(startCol, colCount)));
			moveTo(startRow, startCol);
		}
		return this;
	}
	/**
	 * 当前位置合并
	 * @param rowCount
	 * @param colCount
	 * @return
	 */
	public Excel merge(int rowCount, int colCount){
		return merge(rowNum - 1, rowCount, colNum, colCount);
	}
	/**
	 * 换行
	 * @return
	 */
	public Excel next(){
		skipRow(0);
		return this;
	}
	/**
	 * 切换sheet
	 * @return
	 */
	public boolean nextSheet(){
		if(sheetIndex < numOfSheets - 1){
			switchSheet(sheetIndex + 1);
			return true;
		}
		return false;
	}
	/**
	 * 切换到指定sheet
	 * @param index
	 */
	public Excel switchSheet(int index){
		if(index < 0 || index >= numOfSheets)throw new ExcelException("无效的sheet索引：" + index);
		this.sheet = wb.getSheetAt(index);
		moveTo(0, 0);
		sheetIndex = index;
		
		return this;
	}
	
	//---------------------部分-------------------------
	/**
	 * 当前sheet
	 * @return
	 */
	public Sheet sheet(){
		return sheet;
	}
	/**
	 * 切换到指定位置
	 * @param row 
	 * @param col 
	 */
	public Excel moveTo(int row, int col){
		cur = sheet.getRow(row);
		if(cur == null)cur = sheet.createRow(row);
		rowNum = ++row;
		colNum = col;
		return this;
	}
	/**
	 * 跳过指定列
	 * @param colCount
	 * @return
	 */
	public Excel skip(int colCount){
		colNum += colCount;
		return this;
	}
	/**
	 * 跳过指定行
	 * @param rowCount
	 * @return
	 */
	public Excel skipRow(int rowCount){
		rowNum += rowCount;
		
		moveTo(rowNum, 0);
		return this;
	}
	/**
	 * 获取指定单元格的数据(该方法不会改变索引值)
	 * @param row 行索引
	 * @param col 列索引
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public <T>T get(int row, int col){
		Row r = sheet.getRow(row);
		if(r != null){
			Cell cell = r.getCell(col);
			if(cell != null){
				return (T)ExcelUtil.val(cell);
			}
		}
		return null;
	}
	/**
	 * 获取指定Cell(该方法不会改变索引值)
	 * @param row
	 * @param col
	 * @return
	 */
	public Cell getCell(int row, int col){
		Row r = sheet.getRow(row);
		if(r != null){
			return r.getCell(col);
		}
		return null;
	}
	/**
	 * 遍历当前sheet
	 * @param resolver
	 * @return
	 */
	public Excel iterate(RowResolver resolver){
		Iterator<Row>it = sheet.rowIterator();
		
		int rowNum = 0;
		while(it.hasNext()){
			resolver.resolve(rowNum++, it.next());
		}
		
		return this;
	}
	/**
	 * 遍历当前sheet
	 * @param resolver
	 * @return
	 */
	public Excel iterate(final CellResolver resolver){
		iterate(new CellRowResolverAdapter(resolver));
		
		return this;
	}
	/**
	 * 遍历所有sheet
	 * @param resolver
	 * @return
	 */
	public Excel iterateAll(RowResolver resolver){
		this.switchSheet(0);
		do{
			this.iterate(resolver);
		}while(nextSheet());
		
		return this;
	}
	/**
	 * 遍历所有sheet
	 * @param resolver
	 * @return
	 */
	public Excel iterateAll(final CellResolver resolver){
		iterate(new CellRowResolverAdapter(resolver));
		
		return this;
	}
	/**
	 * 保存当前excel(调用此方法后对象将不可用)
	 */
	public void save(){
		if(this.excelFile == null)throw new ExcelException("目标不存在，请调用save(OutputStream)方法！");
		try {
			save(new FileOutputStream(excelFile));
		} catch (FileNotFoundException e) {
			throw new ExcelException(e);
		}
	}
	/**
	 * 输出到流(调用此方法后对象将不可用)
	 * @param os
	 */
	public void save(OutputStream os){
		try {
			wb.write(os);
		} catch (IOException e) {
			throw new ExcelException(e);
		} finally {
			CloseableUtil.close(this, os);
		}
	}
	/**
	 * 释放资源
	 */
	@Override
	public void close() throws IOException {
		CloseableUtil.close(wb);
		wb = null;
		sheet = null;
		cur = null;
		excelFile = null;
	}
	//合并单元格计算结束列索引
	private int calc(int start, int count){
		return count <= 1 ? start : (start + count - 1);
	}
	//计算列宽度,setWidths方法使用
	private int calcWidth(int width){
		if(width > 255)width = 255;
		return width * 256;
	}
	//转换异常
	private static InputStream getFileInputStream(File file){
		try {
			return new FileInputStream(file);
		} catch (FileNotFoundException e) {
			throw new ExcelException(e);
		}
	}
}
