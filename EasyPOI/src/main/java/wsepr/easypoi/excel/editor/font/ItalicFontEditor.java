package wsepr.easypoi.excel.editor.font;

import wsepr.easypoi.excel.editor.IFontEditor;
import wsepr.easypoi.excel.style.font.Font;

/**
 * 实现一些常用的字体<br/>
 * 该类用于设置斜体
 * @author lxp
 *
 */
public class ItalicFontEditor implements IFontEditor {

	public void updateFont(Font font) {
		font.italic(true);
	}

}
