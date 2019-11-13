package com.example.demo.freemarker;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.IOUtils;

import freemarker.core.ParseException;
import freemarker.template.Configuration;
import freemarker.template.MalformedTemplateNameException;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import freemarker.template.TemplateExceptionHandler;
import freemarker.template.TemplateNotFoundException;

public class ExportDocUtil {

	private Configuration configuration = null;

	private String encoding;

	public ExportDocUtil(String encoding) {
		this.encoding = encoding;
		configuration = new Configuration(Configuration.VERSION_2_3_27);
		configuration.setDefaultEncoding(encoding);
		try {
			configuration.setDirectoryForTemplateLoading(new File("/where/you/store/templates"));
			configuration.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
			configuration.setLogTemplateExceptions(false);
			configuration.setWrapUncheckedExceptions(true);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public Template getTemplate(String templateName) {
		Template template = null;
		try {
			template = configuration.getTemplate(templateName);
		} catch (TemplateNotFoundException e) {
			e.printStackTrace();
		} catch (MalformedTemplateNameException e) {
			e.printStackTrace();
		} catch (ParseException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return template;
	}

	public static String getImageStr(String imagePath) {
		FileInputStream inputStream = null;
		try {
			inputStream = new FileInputStream(imagePath);
			byte[] data = new byte[inputStream.available()];
			inputStream.read(data);
			return data != null ? Base64.encodeBase64String(data) : "";
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(inputStream);
		}
		return "";
	}

	public void exportDoc(String filePath, String templateName, Object data) {
		FileOutputStream fileOutputStream = null;
		OutputStreamWriter outputStreamWriter = null;
		BufferedWriter writer = null;
		try {
			fileOutputStream = new FileOutputStream(filePath);
			outputStreamWriter = new OutputStreamWriter(fileOutputStream, encoding);
			writer = new BufferedWriter(outputStreamWriter);
			getTemplate(templateName).process(data, writer);
		} catch (TemplateException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(fileOutputStream);
			IOUtils.closeQuietly(outputStreamWriter);
			IOUtils.closeQuietly(writer);
		}
	}
}
