package com.example.demo.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class PoiTest {

	public static void main(String[] args) {
		test();
	}

	private static void writeDoc() {
		XWPFDocument doc = PoiUtil.createDocument();
		PoiUtil.createHeading1(doc, "1 样品简介");
		PoiUtil.createHeading2(doc, "1.1 概述");

		PoiUtil.createBody(doc, "asdfasdfsf啊但是发射点VS大哥飞洒地方撒旦飞洒地方嘎嘎发打撒大厦是个大帅哥夫人特温柔各方");

		PoiUtil.createHeading2(doc, "1.2 主要技术参数");
		PoiUtil.createBody(doc, "PICS 表格中用到的缩略语:");
		PoiUtil.createBody(doc, "m : 要求强制支持");
		PoiUtil.createBody(doc, "n/a : 此项不可用");
		PoiUtil.createBody(doc, "o : 可选支持");
		PoiUtil.createBody(doc, "c : 此项是有条件的");
		PoiUtil.createBody(doc, "d : 默认");
		PoiUtil.createBody(doc, "Y : 是");
		PoiUtil.createBody(doc, "N : 否");
		PoiUtil.createTable(doc, "表1.2-1 PISC标识");

		FileOutputStream out = null;
		try {
			out = new FileOutputStream(new File("e:\\word.doc"));
			doc.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void writeDoc2() {

		try {

			XWPFDocument doc = PoiUtil.createDocument("e:\\word.doc");
			doc.createTOC();
			PoiUtil.createHeading1(doc, "1 样品简介");
			PoiUtil.createHeading2(doc, "1.1 概述");

			PoiUtil.createBody(doc, "asdfasdfsf啊但是发射点VS大哥飞洒地方撒旦飞洒地方嘎嘎发打撒大厦是个大帅哥夫人特温柔各方");

			FileOutputStream out = null;
			out = new FileOutputStream(new File("e:\\word2.doc"));
			doc.write(out);
			out.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void test() {
		FileReader fileReader = null;
		FileWriter writer = null;
		try {
			fileReader = new FileReader(new File("e:\\word.doc"));

			// 打开一个写文件器，构造函数中的第二个参数true表示以追加形式写文件
			writer = new FileWriter("e:\\word2.doc", true);
			char[] bytes = new char[2048];
			// 接受读取的内容(n就代表的相关数据，只不过是数字的形式)
			int n = -1;
			// 循环取出数据
			while ((n = fileReader.read(bytes, 0, bytes.length)) != -1) {
				// 写入相关文件
				writer.write(bytes, 0, n);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fileReader != null) {
					fileReader.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
			try {
				if (writer != null) {
					writer.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}

		}

	}

}
