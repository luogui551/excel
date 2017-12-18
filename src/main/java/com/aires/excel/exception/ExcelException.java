package com.aires.excel.exception;
/**
 * @author aires
 * 2017年12月13日 下午2:36:22
 * 描述：
 */
public class ExcelException extends RuntimeException {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	
	public ExcelException(){
		super();
	}
	public ExcelException(String msg){
		super(msg);
	}
	public ExcelException(Throwable cause){
		super(cause);
	}
}
