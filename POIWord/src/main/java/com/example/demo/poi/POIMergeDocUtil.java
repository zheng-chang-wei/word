package com.example.demo.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

/**
 * 合并两个docx文档方法，对文档包含的图片无效
 */
public class POIMergeDocUtil {

	public static void main(String[] args) throws Exception {

		String[] srcDocxs = { "e:\\word.doc", "e:\\word2.doc" };
		String destDocx = "e:\\word3.doc";
		mergeDoc(srcDocxs, destDocx);
	}

	/**
	 * 合并docx文件
	 * 
	 * @param srcDocxs
	 *            需要合并的目标docx文件
	 * @param destDocx
	 *            合并后的docx输出文件
	 */
	public static void mergeDoc(String[] srcDocxs, String destDocx) {

		OutputStream dest = null;
		List<OPCPackage> opcpList = new ArrayList<OPCPackage>();
		int length = null == srcDocxs ? 0 : srcDocxs.length;
		/**
		 * 循环获取每个docx文件的OPCPackage对象
		 */
		for (int i = 0; i < length; i++) {
			String doc = srcDocxs[i];
			OPCPackage srcPackage = null;
			try {
				srcPackage = OPCPackage.open(doc);
			} catch (Exception e) {
				e.printStackTrace();
			}
			if (null != srcPackage) {
				opcpList.add(srcPackage);
			}
		}

		int opcpSize = opcpList.size();
		// 获取的OPCPackage对象大于0时，执行合并操作
		if (opcpSize > 0) {
			try {
				dest = new FileOutputStream(destDocx);
				XWPFDocument src1Document = new XWPFDocument(opcpList.get(0));
				CTBody src1Body = src1Document.getDocument().getBody();
				// OPCPackage大于1的部分执行合并操作
				if (opcpSize > 1) {
					for (int i = 1; i < opcpSize; i++) {
						OPCPackage src2Package = opcpList.get(i);
						XWPFDocument src2Document = new XWPFDocument(src2Package);
						CTBody src2Body = src2Document.getDocument().getBody();
						appendBody(src1Body, src2Body);
					}
				}
				// 将合并的文档写入目标文件中
				src1Document.write(dest);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				// 关闭流
				IOUtils.closeQuietly(dest);
			}
		}

	}

	/**
	 * 合并文档内容
	 * 
	 * @param src
	 *            目标文档
	 * @param append
	 *            要合并的文档
	 * @throws Exception
	 */
	private static void appendBody(CTBody src, CTBody append) throws Exception {
		XmlOptions optionsOuter = new XmlOptions();
		optionsOuter.setSaveOuter();
		String appendString = append.xmlText(optionsOuter);
		String srcString = src.xmlText();
		String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
		String mainPart = srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
		String sufix = srcString.substring(srcString.lastIndexOf("<"));
		String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
		CTBody makeBody = CTBody.Factory.parse(prefix + mainPart + addPart + sufix);
		src.set(makeBody);
	}
}