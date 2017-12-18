package com.aires.test;

import java.util.Arrays;
import java.util.Date;

import org.junit.Test;

import com.aires.excel.Excel;
import com.aires.excel.resolver.SimpleRowResolver;
import com.aires.excel.util.ExcelUtil;

/**
 * @author aires
 * 2017年11月20日 上午11:14:59
 * 描述：
 */
public class TestCase {
	
	private String excelFile = "C:\\Users\\aires\\Desktop\\test.xlsx";
	
	@Test
	public void read(){
		//1、读取测试
		Excel excel = ExcelUtil.parse(excelFile).iterateAll(new SimpleRowResolver() {
			
			@Override
			public void resolve(int rowNum, Object[]val) {
				//控制台输出第一行的内容
				System.out.println(Arrays.toString(val));	
			}			
		});
		System.out.println((Object)excel.switchSheet(0).get(19, 14));
	}
	@Test
	public void write(){
		//2、编辑
		Excel excel = ExcelUtil.parse(excelFile);
		excel.moveTo(10, 3);
		
		excel.write("1", 1, "aadqqf", new Date());
		
		excel.newSheet("测试页").write("1", 1, "aadqqf", new Date()).save();
	}
	@Test
	public void merge(){
		Excel excel = ExcelUtil.parse(excelFile);
		excel.merge(5, 5).write(new Date()).merge(10, 10, 10, 5).write(new Date()).save();
	}
	@Test
	public void skip(){
		Excel excel = ExcelUtil.parse(excelFile);
		excel.skipRow(5).skip(5);
		
		excel.write("第6行,第6列");
	}
	@Test
	public void setWidths(){
		Excel excel = ExcelUtil.parse(excelFile).newSheet("widths");
		excel.setWidths(10);//第一列宽度
		
		excel.setWidths(1, 20, 30, 40, 50);
		
		excel.save();
	}
	
	public static void main(String[] args) {
		new TestCase().read();
	}
}
