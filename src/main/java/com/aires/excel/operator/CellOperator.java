package com.aires.excel.operator;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import com.aires.excel.exception.ExcelException;

/**
 * @author aires
 * 2017年12月13日 下午2:50:24
 * 描述： 支持单元格数据读取或者修改
 */
public abstract class CellOperator {
	
	private static final Map<Class<?>, CellOperator>operatorsMap = new HashMap<Class<?>, CellOperator>();
	private static final Map<CellType, CellOperator>cellTypeMap = new HashMap<CellType, CellOperator>();
	static {
		registe(CellType.STRING, new StringOperator(), String.class);
		registe(CellType.BOOLEAN, new BooleanOperator(), Boolean.class);
		registe(CellType.NUMERIC, new NumbericOperator(), Number.class, Date.class);
		registe(CellType.FORMULA, new FormulaOperator());
	}
	/**
	 * 注册
	 * @param clzz 写入数据时使用
	 * @param cellType 读取数据时使用
	 * @param operator
	 */
	public static void registe(CellType cellType, CellOperator operator, Class<?>...clzz){
		for(Class<?>clz : clzz)
			operatorsMap.put(clz, operator);
		if(cellType != null)cellTypeMap.put(cellType, operator);
	}
	/**
	 * 根据单元格类型获取
	 * @param type
	 * @return
	 */
	public static CellOperator getByType(CellType type){
		return getWithDefault(cellTypeMap.get(type));
	}
	/**
	 * 根据数据类型获取
	 * @param value
	 * @return
	 */
	public static CellOperator getByValue(Object value){
		Class<?>clz = value.getClass();
		if(Number.class.isAssignableFrom(clz))clz = Number.class;
		if(value instanceof String && ((String)value).startsWith("="))return cellTypeMap.get(CellType.FORMULA);
			
		return getWithDefault(operatorsMap.get(clz));
	}
	/**
	 * 根据单元格类型获取，同时设置当前单元格
	 * @param cell
	 * @return
	 */
	public static CellOperator getByCell(Cell cell){
		return getWithDefault(cellTypeMap.get(cell.getCellTypeEnum())).setCell(cell);
	}
	/**
	 * 写入数据
	 * @param cell
	 * @param v
	 */
	public static void write(Cell cell, Object v){
		getByValue(v).setCell(cell).write(v);
	}
	/**
	 * 默认值
	 * @param operator
	 * @return
	 */
	private static CellOperator getWithDefault(CellOperator operator){
		return operator == null ? cellTypeMap.get(CellType.STRING) : operator;
	}
	
	/**
	 * 获取单元格数据
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public <T>T get(){
		return (T)get0();
	}
	
	private static final ThreadLocal<Cell>tl = new ThreadLocal<Cell>();
	/**
	 * 设置待处理单元格
	 * @param cell
	 */
	public CellOperator setCell(Cell cell){
		tl.set(cell);
		return this;
	}
	/**
	 * 当前单元格
	 * @return
	 */
	public Cell getCell(){
		Cell cell = tl.get();
		if(cell == null)throw new ExcelException("请先调用setCell方法");
		return cell;
	}
	
	/**
	 * 获取单元格数据
	 * @return
	 */
	protected abstract Object get0();
	/**
	 * 数据写入单元格
	 * @param v
	 */
	public abstract void write(Object v);
}
