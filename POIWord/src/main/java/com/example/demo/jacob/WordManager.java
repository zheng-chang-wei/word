/**
 * Copyright (C), 2015-2017, XXX有限公司
 * FileName: WordManager
 * Author:   Administrator
 * Date:     2017/8/16 8:50
 * Description:
 * History:
 * <author>          <time>          <version>          <desc>
 * 作者姓名           修改时间           版本号              描述
 */
package com.example.demo.jacob;

/**
 * Created by Administrator on 2017/8/16.
 *
 * @version V1.0
 */

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * 〈一句话功能简述〉<br>
 * 〈〉
 * 
 * @author Administrator
 * @create 2017/8/16
 * @since 1.0.0
 */
public class WordManager {

	// word文档
	private Dispatch doc;

	// word运行程序对象
	private ActiveXComponent word;

	// 所有word文档集合
	private Dispatch documents;

	// 选定的范围或插入点
	private Dispatch selection;

	private boolean saveOnExit = true;

	/** */
	/**
	 * @param visible
	 *            为true表示word应用程序可见
	 */
	public WordManager(boolean visible) {
		if (word == null) {
			word = new ActiveXComponent("Word.Application");
			word.setProperty("Visible", new Variant(visible));
		}
		if (documents == null)
			documents = word.getProperty("Documents").toDispatch();
	}

	/** */
	/**
	 * 设置退出时参数
	 * 
	 * @param saveOnExit
	 *            boolean true-退出时保存文件，false-退出时不保存文件
	 */
	public void setSaveOnExit(boolean saveOnExit) {
		this.saveOnExit = saveOnExit;
	}

	/** */
	/**
	 * 打开一个已存在的文档
	 * 
	 * @param docPath
	 */
	public void openDocument(String docPath) {
		closeDocument();
		doc = Dispatch.call(documents, "Open", docPath).toDispatch();
		selection = Dispatch.get(word, "Selection").toDispatch();
	}

	/************************************** 书签操作 *************************************/
	/** */
	/**
	 * 获取书签个数
	 */
	public int getLabelCount() {
		Dispatch bookMarks = Dispatch.call(doc, "Bookmarks").toDispatch();
		int labelCount = Dispatch.get(bookMarks, "Count").getInt(); // 书签数
		return labelCount;
	}

	/** */
	/**
	 * 获取书签标题
	 * 
	 * @param labelIndex
	 */
	public String getLabelName(int labelIndex) {
		Dispatch bookMarks = Dispatch.call(doc, "Bookmarks").toDispatch();
		Dispatch rangeItem = Dispatch.call(bookMarks, "Item", labelIndex).toDispatch();
		String labelName = Dispatch.call(rangeItem, "Name").toString();
		return labelName;
	}

	/** */
	/**
	 * 获取书签内容
	 * 
	 * @param labelName
	 */
	public String getLabelValue(String labelName) {
		String rawLabelValue = getRawLabelValue(labelName);
		String labelValue = rawLabelValue.replaceAll("", "");
		// 如果含有抬头
		if (labelName.startsWith("head_")) {
			labelValue = rawLabelValue.substring(rawLabelValue.indexOf("：") + 1).replaceAll(" ", "");
		}
		// 如果含有头尾
		if (labelName.startsWith("full_")) {
			labelValue = rawLabelValue.substring(1, rawLabelValue.length() - 1).replaceAll(" ", "");
		}
		if (labelName.startsWith("t_")) {
			labelValue = rawLabelValue.substring(0, rawLabelValue.length() - 2);
		}
		return labelValue;
	}

	/** */
	/**
	 * 获取书签原始内容
	 * 
	 * @param labelName
	 */
	public String getRawLabelValue(String labelName) {
		Dispatch bookMarks = Dispatch.call(doc, "Bookmarks").toDispatch();
		Dispatch rangeItem = Dispatch.call(bookMarks, "Item", labelName).toDispatch();
		Dispatch range = Dispatch.call(rangeItem, "Range").toDispatch();
		String rawLabelValue = Dispatch.get(range, "Text").toString();
		return rawLabelValue;
	}

	/** */
	/**
	 * 设置书签内容
	 * 
	 * @param labelName
	 * @param labelValue
	 */
	public void setLabelValue(String labelName, String labelValue) {
		String rawLabelValue = getRawLabelValue(labelName);
		String toFindString = getLabelValue(labelName);
		System.out.println(toFindString);
		Dispatch bookMarks = Dispatch.call(doc, "Bookmarks").toDispatch();
		Dispatch rangeItem = Dispatch.call(bookMarks, "Item", labelName).toDispatch();
		Dispatch range = Dispatch.call(rangeItem, "Range").toDispatch();
		Dispatch.call(range, "Select");
		// 如果是表格内的标签
		if (labelName.startsWith("t_")) {
			Dispatch.put(selection, "Text", new Variant(labelValue));
		}
		// 如果是表格外的标签
		else {
			if (toFindString == null || toFindString.equals("")) {
				int rawLen = rawLabelValue.length();
				if (labelName.startsWith("head_")) {
					rawLen = rawLabelValue.substring(rawLabelValue.indexOf("：")).length();
				}
				if (labelName.startsWith("full_")) {
					rawLen = rawLabelValue.length() - 2;
				}
				for (int i = 1; i <= rawLen; i++) {
					toFindString += " ";
				}
			}
			replaceText(toFindString, labelValue);
		}

	}

	/** */
	/**
	 * 把选定选定内容设定为替换文本
	 * 
	 * @param toFindString
	 *            查找字符串
	 * @param newString
	 *            要替换的内容
	 * @return
	 */
	public boolean replaceText(String toFindString, String newString) {
		if (!find(toFindString)) {
			System.out.println(toFindString);
			return false;
		}
		// find(toFindString);
		Dispatch.put(selection, "Text", newString);
		return true;
	}

	/************************************** 表操作 ***************************************/
	/** */
	/**
	 * 获取表格行数
	 * 
	 * @param tableIndex
	 *            word文档中的第N张表(从1开始)
	 */
	public int getRowCount(int tableIndex) {
		// 所有表格
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();

		return Dispatch.get(rows, "Count").getInt();
	}

	/** */
	/**
	 * 获取表格某一行的列数
	 * 
	 * @param tableIndex
	 * @param rowIndex
	 */
	public int getRowColumnCount(int tableIndex, int rowIndex) {
		// 所有表格
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		// 所需行
		Dispatch row = Dispatch.call(rows, "Item", new Variant(rowIndex)).toDispatch();
		// 表格数
		Dispatch cells = Dispatch.get(row, "Cells").toDispatch();
		return Dispatch.get(cells, "Count").getInt();
	}

	/** */
	/**
	 * 获取指定的单元格的数据
	 * 
	 * @param tableIndex
	 * @param cellRowIdx
	 * @param cellColIdx
	 */
	public String getCellValue(int tableIndex, int cellRowIdx, int cellColIdx) {
		// 所有表格
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// 要获取的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx), new Variant(cellColIdx)).toDispatch();
		Dispatch range = Dispatch.get(cell, "Range").toDispatch();

		String cellValue = Dispatch.get(range, "Text").toString().replaceAll("", "").trim();

		return cellValue;
	}

	/** */
	/**
	 * 在指定的单元格里填写数据
	 * 
	 * @param tableIndex
	 * @param cellRowIdx
	 * @param cellColIdx
	 * @param cellValue
	 */
	public void setCellValue(int tableIndex, int cellRowIdx, int cellColIdx, String cellValue) {
		// 所有表格
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx), new Variant(cellColIdx)).toDispatch();
		Dispatch.call(cell, "Select");
		Dispatch.put(selection, "Text", cellValue);
	}

	/** */
	/**
	 * 增加一行
	 * 
	 * @param tableIndex
	 *            word文档中的第N张表(从1开始)
	 */
	public void addRow(int tableIndex) {
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch.call(rows, "Add");
	}

	/** */
	/**
	 * 从选定内容或插入点开始查找文本
	 * 
	 * @param toFindString
	 *            要查找的文本
	 * @return boolean true-查找到并选中该文本，false-未查找到文本
	 */
	public boolean find(String toFindString) {
		if (toFindString == null || toFindString.equals(""))
			return false;
		// 从selection所在位置开始查询
		Dispatch find = Dispatch.call(selection, "Find").toDispatch();
		// 设置要查找的内容
		Dispatch.put(find, "Text", toFindString);
		// 向前查找
		Dispatch.put(find, "Forward", "True");
		// 设置格式
		Dispatch.put(find, "Format", "True");
		// 大小写匹配
		Dispatch.put(find, "MatchCase", "True");
		// 全字匹配
		Dispatch.put(find, "MatchWholeWord", "True");
		// 查找并选中
		return Dispatch.call(find, "Execute").getBoolean();
	}

	/** */
	/**
	 * 文件保存或另存为
	 * 
	 * @param savePath
	 *            保存或另存为路径
	 */
	public void save(String savePath) {
		if (savePath.endsWith("doc")) {
			Dispatch.call((Dispatch) Dispatch.call(word, "WordBasic").getDispatch(), "FileSaveAs", savePath, new Variant(24));
		} else {
			int type = 12;
			Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { savePath, new Variant(type) }, new int[1]);
		}

	}

	/** */
	/**
	 * 关闭当前word文档
	 */
	public void closeDocument() {
		if (doc != null) {
			Dispatch.call(doc, "Save");
			Dispatch.call(doc, "Close", new Variant(saveOnExit));
			doc = null;
		}
	}

	/** */
	/**
	 * 关闭全部应用
	 */
	public void close() {
		closeDocument();
		if (word != null) {
			Dispatch.call(word, "Quit");
			word = null;
		}
		selection = null;
		documents = null;
	}

	/** */
	/**
	 * 打印当前word文档
	 */
	public void printFile() {
		if (doc != null) {
			Dispatch.call(doc, "PrintOut");
		}
	}

	/** */
	/**
	 * 在当前插入点插入字符串
	 * 
	 * @param newText
	 *            要插入的新字符串
	 */
	public void insertText(String newText) {
		Dispatch.put(selection, "Text", newText);
	}

	public void insertNewParagraph() {
		Dispatch.call(selection, "EndKey", "5");
		Dispatch.call(selection, "TypeParagraph");
	}

	/**
	 * 换行
	 */
	public void enter() {
		Dispatch.call(selection, "TypeParagraph");

	}

	/** */
	/**
	 * 把插入点移动到文件首位置
	 */
	public void moveStart() {
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		Dispatch.call(selection, "HomeKey", new Variant(6));
	}

	/** */
	/**
	 * 把选定的内容或插入点向上移动
	 * 
	 * @param pos
	 *            移动的距离
	 */
	public void moveUp(int pos) {
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveUp");

	}

	/** */
	/**
	 * 把选定的内容或者插入点向下移动
	 * 
	 * @param pos
	 *            移动的距离
	 */
	public void moveDown(int pos) {
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveDown");
	}

	/**
	 * 把选定的内容或者插入点向右移动
	 * 
	 * @param pos
	 *            移动的距离
	 */
	public void moveLeft(int pos) {
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveLeft");
	}

	/**
	 * 把选定的内容或者插入点向右移动
	 * 
	 * @param pos
	 *            移动的距离
	 */
	public void moveRight(int pos) {
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveRight");
	}

	public void insertToc() {

		// boolean b = find("{目录}");
		//
		// if (b == true) {
		// Dispatch.call(this.selection, "Delete");
		// Dispatch tablesOfContents = Dispatch.call(doc, "TablesOfContents").toDispatch();// 整个目录区域
		//
		// // 整个目录
		// Dispatch tableOfContents = Dispatch.call(tablesOfContents, "Item", new Variant(1)).toDispatch();
		//
		// // 拿到整个目录的范围
		// Dispatch tableOfContentsRange = Dispatch.get(tableOfContents, "Range").toDispatch();
		//
		// Dispatch.call(tableOfContentsRange, "Delete");
		// }
		//
		// moveStart();
		//
		// Dispatch.call(selection, "TypeParagraph");
		//
		// moveDown(3);
		insertText("目  录");
		Dispatch alignment = Dispatch.get(selection, "ParagraphFormat").toDispatch(); // 行列格式化需要的对象
		Dispatch.put(alignment, "Alignment", "1"); // (1:置中 2:靠右 3:靠左)

		moveRight(1);
		// 返回一个**Range** 对象, 该对象代表指定对象中包含的文档部分
		Dispatch range = Dispatch.get(this.selection, "RANGE").toDispatch();

		// 返回一个只读的**fields** 集合, 该集合代表选定内容中的所有域。
		Dispatch fields = Dispatch.call(this.selection, "FIELDS").toDispatch();

		Dispatch.call(fields, "ADD", range, new Variant(-1), new Variant("TOC"), new Variant(true));

		// 返回一个**TablesOfContents** 集合, 该集合代表指定文档中的目录表。 此为只读属性。
		Dispatch tablesOfContents = Dispatch.call(doc, "TablesOfContents").toDispatch();// 整个目录区域

		// 整个目录
		Dispatch tableOfContents = Dispatch.call(tablesOfContents, "Item", new Variant(1)).toDispatch();

		// 拿到整个目录的范围,返回一个Range对象, 该对象代表指定目录中包含的文档部分。
		Dispatch tableOfContentsRange = Dispatch.get(tableOfContents, "Range").toDispatch();
		// // 取消选中,应该就是移动光标
		Dispatch format = Dispatch.get(tableOfContentsRange, "ParagraphFormat").toDispatch();
		// // 设置段落格式为首行缩进2个字符
		Dispatch.put(format, "CharacterUnitLeftIndent", new Variant(1));
		insertNewParagraph();
	}

}
