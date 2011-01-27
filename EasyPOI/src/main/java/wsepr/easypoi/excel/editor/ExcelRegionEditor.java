package wsepr.easypoi.excel.editor;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;

import wsepr.easypoi.excel.ExcelContext;


public class ExcelRegionEditor extends AbstractRegionEditor<ExcelRegionEditor> {

	private CellRangeAddress cellRange;

	public ExcelRegionEditor(int beginRow, int beginCol, int endRow, int endCol, ExcelContext context) {
		super(context);
		cellRange = new CellRangeAddress(beginRow, endRow, beginCol, endCol);
	}
	
	/**
	 * 插入一张图片
	 * @return
	 */
	public ExcelRegionEditor image(String imgPath) {
		ByteArrayOutputStream byteArrayOut = null;
		BufferedImage bufferImg = null;
		try {
			if (imgPath.startsWith("http")) {
				URL url = new URL(imgPath);
				URLConnection conn = url.openConnection();
				bufferImg = ImageIO.read(conn.getInputStream());
			} else {
				bufferImg = ImageIO.read(new File(imgPath));
			}
			HSSFClientAnchor anchor = new HSSFClientAnchor(10, 10, 0, 0, 
					(short) cellRange.getFirstColumn(), cellRange.getFirstRow(), 
					(short) (cellRange.getLastColumn()+1), cellRange.getLastRow()+1);
			anchor.setAnchorType(3);
			HSSFPatriarch patr = ctx.getHSSFPatriarch(this.workingSheet);
			byteArrayOut = new ByteArrayOutputStream();
			ImageIO.write(bufferImg, "JPEG", byteArrayOut);
			int imageIndex = workBook.addPicture(byteArrayOut.toByteArray(),HSSFWorkbook.PICTURE_TYPE_JPEG);
			patr.createPicture(anchor, imageIndex);
		} catch (IOException e) {
			e.printStackTrace();
		} catch(Exception e){
			e.printStackTrace();
		} finally {
			try {
				byteArrayOut.close();
			} catch (Exception e) {
			}
		}
		return this;
	}
	
	/**
	 * 新建一个单元格编辑器，包含所有单元格
	 * @return
	 */
	@Override
	protected ExcelCellEditor newCellEditor(){
		ExcelCellEditor cellEditor = new ExcelCellEditor(this.ctx);
		for(int i=cellRange.getFirstRow(); i<=cellRange.getLastRow() ;i++){
			for(int j=cellRange.getFirstColumn();j<=cellRange.getLastColumn();j++){
				cellEditor.add(i, j);
			}
		}
		return cellEditor;
	}

	@Override
	protected ExcelCellEditor newBottomCellEditor() {
		//下边框
		ExcelCellEditor cellEditorBottom = new ExcelCellEditor(this.ctx);
		for(int i=cellRange.getFirstColumn();i<=cellRange.getLastColumn();i++){
			cellEditorBottom.add(cellRange.getLastRow(), i);
		}
		return cellEditorBottom;
	}

	@Override
	protected ExcelCellEditor newLeftCellEditor() {
		//左边框
		ExcelCellEditor cellEditorLeft = new ExcelCellEditor(this.ctx);
		for(int i=cellRange.getFirstRow();i<=cellRange.getLastRow();i++){
			cellEditorLeft.add(i, cellRange.getFirstColumn());
		}
		return cellEditorLeft;
	}

	@Override
	protected ExcelCellEditor newRightCellEditor() {
		//右边框
		ExcelCellEditor cellEditorRight = new ExcelCellEditor(this.ctx);
		for(int i=cellRange.getFirstRow();i<=cellRange.getLastRow();i++){
			cellEditorRight.add(i, cellRange.getLastColumn());
		}
		return cellEditorRight;
	}

	@Override
	protected ExcelCellEditor newTopCellEditor() {
		//上边框
		ExcelCellEditor cellEditorTop = new ExcelCellEditor(this.ctx);
		for(int i=cellRange.getFirstColumn();i<=cellRange.getLastColumn();i++){
			cellEditorTop.add(cellRange.getFirstRow(), i);
		}
		return cellEditorTop;
	}

	@Override
	protected List<CellRangeAddress> getCellRange() {
		List<CellRangeAddress> cellRangeList = new ArrayList<CellRangeAddress>();
		cellRangeList.add(this.cellRange);
		return cellRangeList;
	}	
}
