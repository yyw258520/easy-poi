package wsepr.easypoi.excel.editor.font;

import wsepr.easypoi.excel.editor.IFontEditor;
import wsepr.easypoi.excel.style.font.BoldWeight;
import wsepr.easypoi.excel.style.font.Font;

/**
 * 实现一些常用的字体<br/>
 * 该类用于把字体加粗
 * @author lxp
 *
 */
public class BoldFontEditor implements IFontEditor {

	public void updateFont(Font font) {
		font.boldweight(BoldWeight.BOLD);
	}

}
