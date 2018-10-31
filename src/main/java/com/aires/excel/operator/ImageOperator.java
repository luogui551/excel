package com.aires.excel.operator;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author aires
 * 2018年1月10日 上午11:39:16
 * 描述：
 */
public class ImageOperator extends CellOperator{

	@Override
	protected Object get0() {
		throw new UnsupportedOperationException("暂不支持图片的读取！");
	}

	@Override
	public void write(Object v) {
		Sheet sheet = getCell().getRow().getSheet();
		Workbook wb = sheet.getWorkbook();
		int pictureIdx = wb.addPicture((byte[])v, Workbook.PICTURE_TYPE_PNG);  
		  
		CreationHelper helper = wb.getCreationHelper();  
		Drawing<?> drawing = sheet.createDrawingPatriarch();  
		ClientAnchor anchor = helper.createClientAnchor();  
		  
		// 图片插入坐标  
		anchor.setCol1(0);  
		anchor.setRow1(1);  
		// 插入图片  
		Picture pict = drawing.createPicture(anchor, pictureIdx);  
		pict.resize(); 
	}

}
