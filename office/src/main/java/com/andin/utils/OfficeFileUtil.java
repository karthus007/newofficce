package com.andin.utils;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.aspose.cells.Border;
import com.aspose.cells.BorderType;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Color;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.slides.Presentation;
import com.aspose.words.Document;

public class OfficeFileUtil {

	private static Logger logger = LoggerFactory.getLogger(OfficeFileUtil.class);
	
	private final static String DOCX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.DOCX_PATH;
	
	private final static String XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.XLSX_PATH;
	
	private final static String PPTX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PPTX_PATH;
	
	//private final static String HTML_DOCX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.HTML_DOCX_PATH;
	
	//private final static String HTML_XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.HTML_XLSX_PATH;
	
	//private final static String HTML_PPTX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.HTML_PPTX_PATH;
	
	private final static String PDF_DOCX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_DOCX_PATH;
	
	private final static String PDF_XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_XLSX_PATH;
	
	private final static String PDF_PPTX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_PPTX_PATH;

	public static void main(String[] args) {
		asposeWordToPdf("c:/app/bug.docx", "c:/app/test.html");
	}
	
	/**
	 * word转pdf
	 * @param inputFileName
	 * @param outputFileName
	 * @throws Exception
	 */
	private static boolean asposeWordToPdf(String inputFileName, String outputFileName){
		boolean result = false;
		try {
			byte[] bytes = ConstantUtil.ASPOSE_WORD_LICENSE.getBytes("UTF-8");
			InputStream in =  new ByteArrayInputStream(bytes);
			com.aspose.words.License asposeLic = new com.aspose.words.License();
			asposeLic.setLicense(in);
			Document convertDoc = new Document(inputFileName);
			if(convertDoc.getProtectionType() == 2) {
				convertDoc.unprotect("http://www.gztemco.com/");
				//convertDoc.acceptAllRevisions();
				//convertDoc.protect(2, "http://www.gztemco.com/");
			}
			convertDoc.acceptAllRevisions();
			convertDoc.save(outputFileName, com.aspose.words.SaveFormat.HTML);
			in.close();
			result = true;
			logger.debug("OfficeFileUtil.asposeWordToPdf method executed is successful, output file path is: " + outputFileName);
		}  catch (Exception e) {
			logger.error("OfficeFileUtil.asposeWordToPdf method executed is error: ", e);
		}
        return result;
	}
	
	/**
	 * excel转pdf
	 * @param inputFileName
	 * @param outputFileName
	 * @throws Exception
	 */
	public static boolean asposeExcelToPdf(String inputFileName, String outputFileName){
		boolean result = false;
		try {
			byte[] bytes = ConstantUtil.ASPOSE_WORD_LICENSE.getBytes("UTF-8");
			InputStream in =  new ByteArrayInputStream(bytes);
			com.aspose.cells.License asposeLic = new com.aspose.cells.License();
			asposeLic.setLicense(in);
       	 	Workbook book = new Workbook(inputFileName);
       	 	book.save(outputFileName, com.aspose.cells.SaveFormat.PDF);
			in.close();
			result = true;
			logger.debug("OfficeFileUtil.asposeExcelToPdf method executed is successful, output file path is: " + outputFileName);
		}  catch (Exception e) {
			logger.error("OfficeFileUtil.asposeExcelToPdf method executed is error: ", e);
		}
        return result;
	}
	
	/**
	 * excel转html
	 * @param inputFileName
	 * @param outputFileName
	 * @throws Exception
	 */
	public static boolean asposeExcelToHtml(String inputFileName, String outputFileName){
		boolean result = false;
		try {
			byte[] bytes = ConstantUtil.ASPOSE_WORD_LICENSE.getBytes("UTF-8");
			InputStream in =  new ByteArrayInputStream(bytes);
			com.aspose.cells.License asposeLic = new com.aspose.cells.License();
			asposeLic.setLicense(in);
       	 	Workbook book = new Workbook(inputFileName);
       	 	Style style = book.createStyle();
			Border top = style.getBorders().getByBorderType(BorderType.TOP_BORDER);
			top.setLineStyle(CellBorderType.THIN);
			top.setColor(Color.fromArgb(211, 211, 211));
			Border bottom = style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER);
			bottom.setLineStyle(CellBorderType.THIN);
			bottom.setColor(Color.fromArgb(211, 211, 211));
			Border left = style.getBorders().getByBorderType(BorderType.LEFT_BORDER);
			left.setLineStyle(CellBorderType.THIN);
			left.setColor(Color.fromArgb(211, 211, 211));
			Border right = style.getBorders().getByBorderType(BorderType.RIGHT_BORDER);
			right.setLineStyle(CellBorderType.THIN);
			right.setColor(Color.fromArgb(211, 211, 211));
			book.setDefaultStyle(style);
       	 	book.save(outputFileName, com.aspose.cells.SaveFormat.HTML);
			in.close();
			result = true;
			logger.debug("OfficeFileUtil.asposeExcelToHtml method executed is successful, output file path is: " + outputFileName);
		}  catch (Exception e) {
			logger.error("OfficeFileUtil.asposeExcelToHtml method executed is error: ", e);
		}
        return result;
	}

	/**
	 * pptx转pdf
	 * @param inputFileName
	 * @param outputFileName
	 * @throws Exception
	 */
	private static boolean asposePptxToPdf(String inputFileName, String outputFileName){
		boolean result = false;
		try {
			byte[] bytes = ConstantUtil.ASPOSE_WORD_LICENSE.getBytes("UTF-8");
			InputStream in =  new ByteArrayInputStream(bytes);
			com.aspose.slides.License asposeLic = new com.aspose.slides.License();
			asposeLic.setLicense(in);
        	Presentation pres = new Presentation(inputFileName);
        	pres.save(outputFileName, com.aspose.slides.SaveFormat.Pdf);
			in.close();
			result = true;
			logger.debug("OfficeFileUtil.asposePptxToPdf method executed is successful, output file path is: " + outputFileName);
		}  catch (Exception e) {
			logger.error("OfficeFileUtil.asposePptxToPdf method executed is error: ", e);
		}
        return result;
	}
	
	/**
	 * HTML转换PDF
	 * @param fileInfo /app/file/filename (不包含后缀.xlsx)
	 * @param inputPath 
	 * @param outputPath
	 * @return
	 * @throws Exception
	 */
	public static boolean htmlToPdf(String fileInfo, String inputPath, String outputPath) throws Exception {
		boolean result = false;
		String htmlFileNamePath = inputPath + fileInfo + ConstantUtil.HTML;	
		String pdfFileNamePath = outputPath + fileInfo + ConstantUtil.PDF;
		result = OfficeCmdUtil.htmlToPdf(htmlFileNamePath, pdfFileNamePath);
		logger.debug("输入文件为：" + fileInfo + ".html, html转pdf的结果为：" + result);
		return result;
	}
	
	public static boolean officeToPdf(String inputFileName) {
		boolean result = false;
		try {
			logger.debug("OfficeFileUtil.officeToPdf 转换的文件名为： " + inputFileName);
			int index = inputFileName.lastIndexOf(".");
			String fileName = inputFileName.substring(0, index);
			String fileType = inputFileName.substring(index);
			if(ConstantUtil.DOCX.equals(fileType) || ConstantUtil.DOC.equals(fileType)) {
				//将DOCX文件转换为PDF
				result = asposeWordToPdf(DOCX_PATH + inputFileName, PDF_DOCX_PATH + fileName + ConstantUtil.PDF);
				//result = OfficeCmdUtil.wordToHtml(DOCX_PATH + inputFileName, HTML_DOCX_PATH);
				logger.debug("输入文件为：" + inputFileName + ", docx转pdf的结果为：" + result);
				FileUtil.deleteFilePath(DOCX_PATH + inputFileName);
				//if(result) {
				//	result = htmlToPdf(fileName, HTML_DOCX_PATH, PDF_DOCX_PATH);
				//	logger.debug("输入文件为：" + inputFileName + ", html转pdf的结果为：" + result);
				//	deleteHtmlFileByName(HTML_DOCX_PATH, fileName);
				//}
				
			}else if(ConstantUtil.XLSX.equals(fileType) || ConstantUtil.XLS.equals(fileType)) {
				//将XLSX文件转换为PDF
				result = asposeExcelToPdf(XLSX_PATH + inputFileName, PDF_XLSX_PATH + fileName + ConstantUtil.PDF);
				//result = OfficeCmdUtil.excelToHtml(XLSX_PATH + inputFileName, HTML_XLSX_PATH + fileName + ConstantUtil.HTML);
				//result = asposeExcelToHtml(XLSX_PATH + inputFileName, HTML_XLSX_PATH + fileName + ConstantUtil.HTML);
				logger.debug("输入文件为：" + inputFileName + ", xlsx转pdf的结果为：" + result);
				FileUtil.deleteFilePath(XLSX_PATH + inputFileName);
				//if(result) {
					//将html文件压缩成zip包
				//	result = FileUtil.getHtmlFileZipByFileName(fileName);
				//	logger.debug("输入文件为：" + inputFileName + ", html文件压缩成zip的结果为：" + result);
				//}
				//if(result) {
					//获取excel每个sheet转成的html文件名列表
				//	List<String> list = getXlsxHtmlFileNameList(fileName);
				//	int size = list.size();
					//将每个html转换成pdf
				//	for (int i = 0; i < size; i++) {
				//		String name = list.get(i);
				//		result = htmlToPdf(name, HTML_XLSX_PATH, PDF_XLSX_PATH);
				//	}
				//	logger.debug("输入文件为：" + inputFileName + ", html转pdf的结果为：" + result);
				//	deleteHtmlFileByName(HTML_XLSX_PATH, fileName);
				//}
			
			}else if(ConstantUtil.PPTX.equals(fileType) || ConstantUtil.PPT.equals(fileType)) {
				//将PPTX文件转换为PDF
				result = asposePptxToPdf(PPTX_PATH + inputFileName, PDF_PPTX_PATH + fileName + ConstantUtil.PDF);
				//result = OfficeCmdUtil.pptToHtml(PPTX_PATH + inputFileName, HTML_PPTX_PATH  + fileName + ConstantUtil.HTML);
				logger.debug("输入文件为：" + inputFileName + ", pptx转pdf的结果为：" + result);
				FileUtil.deleteFilePath(PPTX_PATH + inputFileName);
				//if(result) {
				//	result = htmlToPdf(fileName, HTML_PPTX_PATH, PDF_PPTX_PATH);
				//	logger.debug("输入文件为：" + inputFileName + ", html转pdf的结果为：" + result);
				//	FileUtils.forceDelete(new File(HTML_PPTX_PATH + fileName + ConstantUtil.HTML));
				//}
			}else {
				logger.error("OfficeFileUtil.officeToPdf 需转换的文件格式不符合规范：" + inputFileName);
				return false;
			}
			logger.debug("OfficeFileUtil.officeToPdf 转换执行成功！ 文件名为：" + inputFileName);
		} catch (Exception e) {
			result = false;
			logger.error("OfficeFileUtil.officeToPdf method executed is failed : ", e);
		}
		return result;
	}
	
	
	/**
	  * 通过文件夹路径和匹配到的文件名删除文件
	 * @param fileDirPath
	 * @param fileName
	 * @return
	 */
	public static boolean deleteHtmlFileByName(String fileDirPath, String fileName) {
		logger.debug("OfficeFileUtil.deleteHtmlFileByName method params is: [fileDirPath=" + fileDirPath + "], [fileName=" + fileName + "]");
		boolean result = false;
		try {
			File dir = new File(fileDirPath);
			File[] files = dir.listFiles();
			if(files != null) {
				for (File file : files) {
					String name = file.getName();
					if(name.startsWith(fileName)) {
						FileUtil.deleteFile(file);
						logger.debug("file delete is successfuled, file name is: " + name);
					}
				}
			}
			result = true;
			logger.debug("OfficeFileUtil.deleteHtmlFileByName method executed is successful...");
		} catch (Exception e) {
			logger.error("OfficeFileUtil.deleteHtmlFileByName method executed is error: ", e);
		}
		return result;
	}
	
}
