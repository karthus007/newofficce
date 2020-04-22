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

public class OfficeUtil {
	
	private static Logger logger = LoggerFactory.getLogger(OfficeFileUtil.class);
	
	// WORD转PDF
	public static final int WORD_FORMAT_PDF = 17;
	// DOC转DOCX
	private static final int DOC_FORMAT_DOCX = 12;
	
	private final static String IMAGE_XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_XLSX_PATH;
	
	/**
	 * EXCEL转png
	 * @param inputFileName d:/app/a.xlsx
	 * @param outputFileName d:/app/a
	 * @param type .png
	 * @return
	 */
	public static boolean asposeExcelToImage(String inputFileName, String fileName, String type){
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
    
    public static void main(String[] args) {
		officeWordToPdf("d:/app/test.docx", "d:/app/test.pdf");
	}

}
