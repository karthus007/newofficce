package com.andin.utils;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;
import java.util.List;

import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.aspose.cells.Border;
import com.aspose.cells.BorderType;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Color;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.slides.Presentation;
import com.aspose.words.Document;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class OfficeFileUtil {

	private static Logger logger = LoggerFactory.getLogger(OfficeFileUtil.class);
	
    private static final String OFFICE_EXCEL_TYPE = PropertiesUtil.getProperties("office.excel.type", null);
    
    private static final String OFFICE_CONVERT_TYPE = PropertiesUtil.getProperties("office.convert.type", null);
	
	// WORD转PDF
	public static final int WORD_FORMAT_PDF = 17;
	// DOC转DOCX
	private static final int DOC_FORMAT_DOCX = 12;
	
	private final static String XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.XLSX_PATH;

	private final static String IMAGE_XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_XLSX_PATH;
	
	private final static String HTML_XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.HTML_XLSX_PATH;
		
	/**
	 * office to pdf
	 * @param inputFileName
	 * @return
	 */
	public static boolean officeToPdf(String inputFileName) {
		long startTime = System.currentTimeMillis();
		boolean result = false;
		try {
			String systemType = StringUtil.getSystemType();
			if(ConstantUtil.WINDOWS.equals(systemType)) {
				result = windowsOfficeToPdf(inputFileName);
			}else {
				result = linuxOfficeToPdf(inputFileName);
			}
			logger.debug("OfficeFileUtil.officeToPdf 转换执行成功！ 文件名为：" + inputFileName);
		} catch (Exception e) {
			result = false;
			logger.error("OfficeFileUtil.officeToPdf method executed is failed : ", e);
		}
		long endTime = System.currentTimeMillis();
	    logger.debug("OfficeFileUtil.officeToPdf method executed spend time is: " + (endTime - startTime)/1000 + "s");
		return result;
	}
	
	/**
	 * linux系统实现office转pdf
	 * @param inputFileName
	 * @return
	 */
	public static boolean linuxOfficeToPdf(String inputFileName) {
		boolean result = false;
		try {
			int repeatCount = 3;
			logger.debug("OfficeFileUtil.linuxOfficeToPdf 转换的文件名为： " + inputFileName);
			int index = inputFileName.lastIndexOf(".");
			// 文件名前缀
			String fileName = inputFileName.substring(0, index);
			// 文件名后缀
			String fileType = inputFileName.substring(index);
			// 获取输入文件路径
			String input = StringUtil.getInputFilePathByFileName(inputFileName);
			if(ConstantUtil.DOCX.equals(fileType) || ConstantUtil.DOC.equals(fileType)) {
				String output = StringUtil.getOutputFilePathByFileName(inputFileName);
				//将DOCX文件转换为PDF
				for (int i = 0; i < repeatCount; i++) {
					//存在输出文件则删除
					FileUtil.deleteFilePath(output);
					//将DOCX文件转换为PDF
					
					// aspose convert to pdf
					if("1".equals(OFFICE_CONVERT_TYPE)) {
						result = asposeWordToPdf(input, output);
					// default libreoffice convert to pdf
					}else {
						result = OfficeCmdUtil.officeToPdf(input, output);
					}
					if(result) {
						break;
					}
					logger.debug("OfficeFileUtil.linuxOfficeToPdf 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
				}
				logger.debug("输入文件为：" + inputFileName + ", docx转pdf的结果为：" + result);
				FileUtil.deleteFilePath(input);
			}else if(ConstantUtil.XLSX.equals(fileType) || ConstantUtil.XLS.equals(fileType)) {
				// excel to html
				if("1".equals(OFFICE_EXCEL_TYPE)) {
					// html文件路径
					String output = HTML_XLSX_PATH + fileName + ConstantUtil.HTML;
					//获取生成的文件名前缀
					String prefix = fileName;
					for (int i = 0; i < repeatCount; i++) {
						//将EXCEL文件转换为HTML
						result = asposeExcelToHtml(input, output);
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.asposeExcelToHtml 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
					}
					logger.debug("输入文件为：" + inputFileName + ", xlsx转html的结果为：" + result);
					if(result) {
						//将html文件压缩成zip包
						result = FileUtil.getFileZipByMatchFileNamePrefix(fileName, HTML_XLSX_PATH, prefix);
						logger.debug("输入文件为：" + inputFileName + ", html文件压缩成zip的结果为：" + result);
					}
				// excel to pdf
				}else if("2".equals(OFFICE_EXCEL_TYPE)) {
					String output = StringUtil.getOutputFilePathByFileName(inputFileName);
					//将PPTX文件转换为PDF
					for (int i = 0; i < repeatCount; i++) {
						//存在输出文件则删除
						FileUtil.deleteFilePath(output);
						//将EXCEL文件转换为PDF
						
						// aspose convert to pdf
						if("1".equals(OFFICE_CONVERT_TYPE)) {
							result = asposeExcelToPdf(input, output);
						// default libreoffice convert to pdf
						}else {
							result = OfficeCmdUtil.officeToPdf(input, output);
						}
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.libreofficeExcelToPdf 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
					}
				// default excel to png
				}else {
					//获取生成的文件名前缀
					String prefix = fileName + "-";
					for (int i = 0; i < repeatCount; i++) {
						//将EXCEL文件转换为PNG
						result = asposeExcelToImage(input, prefix, ConstantUtil.PNG);
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.asposeExcelToImage 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + XLSX_PATH + inputFileName);
					}
					logger.debug("输入文件为：" + inputFileName + ", xlsx转png的结果为：" + result);
					FileUtil.deleteFilePath(input);
					if(result) {
						//将png文件压缩成zip包
						result = FileUtil.getFileZipByMatchFileNamePrefix(fileName, IMAGE_XLSX_PATH, prefix);
						logger.debug("输入文件为：" + inputFileName + ", png文件压缩成zip的结果为：" + result);
					}
				}
				FileUtil.deleteFilePath(input);
			}else if(ConstantUtil.PPTX.equals(fileType) || ConstantUtil.PPT.equals(fileType)) {
				String output = StringUtil.getOutputFilePathByFileName(inputFileName);
				//将PPTX文件转换为PDF
				for (int i = 0; i < repeatCount; i++) {
					//存在输出文件则删除
					FileUtil.deleteFilePath(output);
					//将PPT文件转换为PDF
					
					// aspose convert to pdf
					if("1".equals(OFFICE_CONVERT_TYPE)) {
						result = asposePptxToPdf(input, output);
					// default libreoffice convert to pdf
					}else {
						result = OfficeCmdUtil.officeToPdf(input, output);
					}
					if(result) {
						break;
					}
					logger.debug("OfficeFileUtil.linuxOfficeToPdf 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
				}
				logger.debug("输入文件为：" + inputFileName + ", pptx转pdf的结果为：" + result);
				FileUtil.deleteFilePath(input);
			}else {
				logger.error("OfficeFileUtil.linuxOfficeToPdf 需转换的文件格式不符合规范：" + inputFileName);
				return false;
			}
			logger.debug("OfficeFileUtil.linuxOfficeToPdf 转换执行成功！ 文件名为：" + inputFileName);
		} catch (Exception e) {
			result = false;
			logger.error("OfficeFileUtil.linuxOfficeToPdf method executed is failed : ", e);
		}
		return result;
	}
	
	/**
	 * windows系统实现office转pdf
	 * @param inputFileName
	 * @return
	 */
	public static boolean windowsOfficeToPdf(String inputFileName) {
		boolean result = false;
		try {
			int repeatCount = 3;
			logger.debug("OfficeFileUtil.windowsOfficeToPdf 转换的文件名为： " + inputFileName);
			int index = inputFileName.lastIndexOf(".");
			// 文件名前缀
			String fileName = inputFileName.substring(0, index);
			// 文件名后缀
			String fileType = inputFileName.substring(index);
			// 获取输入文件路径
			String input = StringUtil.getInputFilePathByFileName(inputFileName);
			if(ConstantUtil.DOCX.equals(fileType) || ConstantUtil.DOC.equals(fileType)) {
				String output = StringUtil.getOutputFilePathByFileName(inputFileName);
				//将DOCX文件转换为PDF
				for (int i = 0; i < repeatCount; i++) {
					//存在输出文件则删除
					FileUtil.deleteFilePath(output);
					
					// aspose convert to pdf
					if("1".equals(OFFICE_CONVERT_TYPE)) {
						result = asposeWordToPdf(input, output);
					// default office convert to pdf
					}else {
						result = officeWordToPdf(input, output);
					}
					if(result) {
						break;
					}
					logger.debug("OfficeFileUtil.windowsOfficeToPdf 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
				}
				logger.debug("输入文件为：" + inputFileName + ", docx转pdf的结果为：" + result);
				FileUtil.deleteFilePath(input);
			}else if(ConstantUtil.XLSX.equals(fileType) || ConstantUtil.XLS.equals(fileType)) {
				// excel to html
				if("1".equals(OFFICE_EXCEL_TYPE)) {
					// html文件路径
					String output = HTML_XLSX_PATH + fileName + ConstantUtil.HTML;
					//获取生成的文件名前缀
					String prefix = fileName;
					for (int i = 0; i < repeatCount; i++) {
						//将EXCEL文件转换为HTML
						result = asposeExcelToHtml(input, output);
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.asposeExcelToHtml 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
					}
					logger.debug("输入文件为：" + inputFileName + ", xlsx转html的结果为：" + result);
					if(result) {
						//将html文件压缩成zip包
						result = FileUtil.getFileZipByMatchFileNamePrefix(fileName, HTML_XLSX_PATH, prefix);
						logger.debug("输入文件为：" + inputFileName + ", html文件压缩成zip的结果为：" + result);
					}
				// excel to pdf
				}else if("2".equals(OFFICE_EXCEL_TYPE)) {
					String output = StringUtil.getOutputFilePathByFileName(inputFileName);
					//将PPTX文件转换为PDF
					for (int i = 0; i < repeatCount; i++) {
						//存在输出文件则删除
						FileUtil.deleteFilePath(output);
						//将EXCEL文件转换为PDF
						result = asposeExcelToPdf(input, output);
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.windowsOfficeToPdf 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
					}
				// default excel to png
				}else {
					//获取生成的文件名前缀
					String prefix = fileName + "-";
					for (int i = 0; i < repeatCount; i++) {
						//将EXCEL文件转换为PNG
						result = asposeExcelToImage(input, prefix, ConstantUtil.PNG);
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.asposeExcelToImage 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + XLSX_PATH + inputFileName);
					}
					logger.debug("输入文件为：" + inputFileName + ", xlsx转png的结果为：" + result);
					FileUtil.deleteFilePath(input);
					if(result) {
						//将png文件压缩成zip包
						result = FileUtil.getFileZipByMatchFileNamePrefix(fileName, IMAGE_XLSX_PATH, prefix);
						logger.debug("输入文件为：" + inputFileName + ", png文件压缩成zip的结果为：" + result);
					}
				}
				FileUtil.deleteFilePath(input);
			}else if(ConstantUtil.PPTX.equals(fileType) || ConstantUtil.PPT.equals(fileType)) {
				String output = StringUtil.getOutputFilePathByFileName(inputFileName);
				//将PPTX文件转换为PDF
				for (int i = 0; i < repeatCount; i++) {
					//存在输出文件则删除
					FileUtil.deleteFilePath(output);
					//将PPT文件转换为PDF
					result = asposePptxToPdf(input, output);
					if(result) {
						break;
					}
					logger.debug("OfficeFileUtil.windowsOfficeToPdf 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
				}
				logger.debug("输入文件为：" + inputFileName + ", pptx转pdf的结果为：" + result);
				FileUtil.deleteFilePath(input);
			}else {
				logger.error("OfficeFileUtil.windowsOfficeToPdf 需转换的文件格式不符合规范：" + inputFileName);
				return false;
			}
			logger.debug("OfficeFileUtil.windowsOfficeToPdf 转换执行成功！ 文件名为：" + inputFileName);
		} catch (Exception e) {
			result = false;
			logger.error("OfficeFileUtil.windowsOfficeToPdf method executed is failed : ", e);
		}
		return result;
	}
	
	
	/**
	 * EXCEL转png
	 * @param inputFileName d:/app/a.xlsx
	 * @param outputFileName d:/app/a
	 * @param type .png
	 * @return
	 */
	private static boolean asposeExcelToImage(String inputFileName, String fileName, String type){
		boolean result = false;
		try {
			String outputFileName = IMAGE_XLSX_PATH + fileName + "-";
			byte[] bytes = ConstantUtil.ASPOSE_WORD_LICENSE.getBytes("UTF-8");
			InputStream in =  new ByteArrayInputStream(bytes);
			com.aspose.cells.License asposeLic = new com.aspose.cells.License();
			asposeLic.setLicense(in);
       	 	Workbook book = new Workbook(inputFileName);
       	 	//设置默认表格样式
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
            WorksheetCollection sheets = book.getWorksheets();
            //设置图片样式
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            ImageFormat format = null;
            if(ConstantUtil.PNG.equals(type)) {
            	format = ImageFormat.getPng();
            }else {
            	format = ImageFormat.getJpeg();
            }
            imgOptions.setImageFormat(format);
            imgOptions.setCellAutoFit(true);
            imgOptions.setOnePagePerSheet(true);
            for (int i = 0, count = sheets.getCount(); i < count; i++) {
                Worksheet sheet = sheets.get(i);
                //sheet.getPageSetup().setLeftMargin(0);
                //sheet.getPageSetup().setRightMargin(0);
                //sheet.getPageSetup().setBottomMargin(0);
                //sheet.getPageSetup().setTopMargin(0);
                SheetRender render = new SheetRender(sheet, imgOptions);
                render.toImage(0,  outputFileName + (i+1) + type);
                
			}
			in.close();
			result = true;
			logger.debug("OfficeFileUtil.asposeExcelToImage method executed is successful, output file path is: " + outputFileName);
		}  catch (Exception e) {
			logger.error("OfficeFileUtil.asposeExcelToImage method executed is error: ", e);
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
	 * word文档接收修订
	 * @param inputFileName
	 * @param outputFileName
	 * @throws Exception
	 */
	public static boolean asposeWordAcceptRevisions(String inputFileName, String outputFileName){
		boolean result = false;
		try {
			byte[] bytes = ConstantUtil.ASPOSE_WORD_LICENSE.getBytes("UTF-8");
			InputStream in =  new ByteArrayInputStream(bytes);
			com.aspose.words.License asposeLic = new com.aspose.words.License();
			asposeLic.setLicense(in);
			Document convertDoc = new Document(inputFileName);
			convertDoc.acceptAllRevisions();
			convertDoc.save(outputFileName);
			in.close();
			result = true;
			logger.debug("OfficeFileUtil.asposeWordAcceptRevisions method executed is successful, output file path is: " + outputFileName);
		}  catch (Exception e) {
			logger.error("OfficeFileUtil.asposeWordAcceptRevisions method executed is error: ", e);
		}
        return result;
	}
	
	/**
	 * word转pdf
	 * @param inputFileName
	 * @param outputFileName
	 * @throws Exception
	 */
	public static boolean asposeWordToPdf(String inputFileName, String outputFileName){
		boolean result = false;
		try {
			byte[] bytes = ConstantUtil.ASPOSE_WORD_LICENSE.getBytes("UTF-8");
			InputStream in =  new ByteArrayInputStream(bytes);
			com.aspose.words.License asposeLic = new com.aspose.words.License();
			asposeLic.setLicense(in);
			Document convertDoc = new Document(inputFileName);
			convertDoc.save(outputFileName, com.aspose.words.SaveFormat.PDF);
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
	 * pptx转pdf
	 * @param inputFileName
	 * @param outputFileName
	 * @throws Exception
	 */
	public static boolean asposePptxToPdf(String inputFileName, String outputFileName){
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
	 * windows调用office将word转pdf
	 * @param inputFileName
	 * @param outputFileName
	 * @return
	 */
	public static boolean officeWordToPdf(String inputFileName,String outputFileName){
		long startTime = System.currentTimeMillis();
		boolean result = false;
		ActiveXComponent app = null;
		Dispatch doc = null;
		try {
			//打开word应用程序
			app = new ActiveXComponent("Word.Application");
			//设置word不可见，否则会弹出word界面
			app.setProperty("Visible", false);
			//获得word中所有打开的文档,返回Documents对象
			Dispatch docs = app.getProperty("Documents").toDispatch();
			//调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
			doc = Dispatch.call(docs, "Open", inputFileName, false, true).toDispatch();
			//调用Document对象的SaveAs方法，将文档保存为pdf格式
			Dispatch.call(doc, "ExportAsFixedFormat", outputFileName, WORD_FORMAT_PDF);
			result = true;
			logger.debug("OfficeFileUtil.officeWordToPdf method executed is successful, output file path is: " + outputFileName);
		} catch (Exception e) {
			logger.error("OfficeFileUtil.officeWordToPdf method executed is error: ", e);
		} finally {
			// Dispatch.call(doc, "Close", false);  
			Dispatch.call(doc, "Close", new Variant(0));  
            if (app != null) {      
            	// app.invoke("Quit", new Variant[] {});
                app.invoke("Quit", new Variant(0));      
            }
            ComThread.Release();
		}
		long endTime = System.currentTimeMillis();
	    logger.debug("OfficeFileUtil.officeWordToPdf method executed spend time is: " + (endTime - startTime)/1000 + "s");
	    return result;
	}
	
	/**
	 * windows调用office将doc转docx
	 * @param inputFileName
	 * @param outputFileName
	 * @return
	 */
	public static boolean officeDocToDocx(String inputFileName,String outputFileName){
		long startTime = System.currentTimeMillis();
		boolean result = false;
		ActiveXComponent app = null;
		Dispatch doc = null;
		try {
			//打开word应用程序
			app = new ActiveXComponent("Word.Application");
			//设置word不可见，否则会弹出word界面
			app.setProperty("Visible", false);
			//获得word中所有打开的文档,返回Documents对象
			Dispatch docs = app.getProperty("Documents").toDispatch();
			//调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
			doc = Dispatch.call(docs, "Open", inputFileName, false, true).toDispatch();
			//调用Document对象的SaveAs方法，将文档保存为pdf格式
			Dispatch.call(doc, "SaveAs", outputFileName, DOC_FORMAT_DOCX);
			result = true;
			logger.debug("OfficeFileUtil.officeDocToDocx method executed is successful, output file path is: " + outputFileName);
		}  catch (Exception e) {
			logger.error("OfficeFileUtil.officeDocToDocx method executed is error: ", e);
		} finally {
			// Dispatch.call(doc, "Close", false);  
			Dispatch.call(doc, "Close", new Variant(0));  
            if (app != null) {      
                // app.invoke("Quit", new Variant[] {});
                app.invoke("Quit", new Variant(0));      
            }
            ComThread.Release();
		}
		long endTime = System.currentTimeMillis();
	    logger.debug("OfficeFileUtil.officeDocToDocx method executed spend time is: " + (endTime - startTime)/1000 + "s");
	    return result;
	}
	
	public static void main(String[] args) {
		//officeDocToDocx("d:/app/test.doc", "d:/app/test.docx");
		officeWordToPdf("d:/app/test.docx", "d:/app/test.pdf");
	}
	
    /**
           * 获取PDF文本内容
     * @param fileNamePath
     * @return
     */
    public static String getPDFText(String fileNamePath){
		logger.debug("OfficeFileUtil.checkPdfPage method executed is start...");
    	String content = "";
		try {
	        File file = new File(fileNamePath);
	        PDFTextStripper stripper = new PDFTextStripper();
	        PDDocument document = PDDocument.load(file);
	        content = stripper.getText(document).trim();
	        document.close();
	        logger.debug("OfficeFileUtil.getPDFText method executed is successful, fileNamePath is: " + fileNamePath);
		} catch (Exception e) {
			logger.error("OfficeFileUtil.getPDFText method executed is error: ", e);
        }
        return content;
    }
    
    /**
          * 拼接PDF
     * @param tempFileName
     * @param outputFileName
     * @return
     */
    public static boolean mergePDFFile(List<String> filePathList, String outputFileName){
		logger.debug("OfficeFileUtil.mergePdfPage method executed is start...");
		boolean result = false;
        try {
            PDFMergerUtility mergePdf = new PDFMergerUtility();  
            for (String filePathName : filePathList) {
                mergePdf.addSource(filePathName);
			}
            //合并生成PDF文件
            mergePdf.setDestinationFileName(outputFileName);  
            mergePdf.mergeDocuments(null);
            result = true;
			logger.debug("OfficeFileUtil.mergePDFFile method executed is successful...");
        } catch (Exception e) {
			logger.error("OfficeFileUtil.mergePDFFile method executed is error: ", e);
        }
        return result;
    }

}
