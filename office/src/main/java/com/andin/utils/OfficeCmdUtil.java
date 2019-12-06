package com.andin.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * office文件转html工具类
 * @author Administrator
 *
 */
public class OfficeCmdUtil {
	
	private static Logger logger = LoggerFactory.getLogger(OfficeCmdUtil.class);
	
    private static final String EXCEL_TYPE = "xlsx";
    
    private static final String PPT_TYPE = "pptx";
	
    private static final String MONO_CMD = "mono WordToHtml.exe ";
    
    private static final String DOTNET_CMD = "dotnet OfficeTool.dll ";
    
    private static final String WKHTMLTOPDF_CMD = "wkhtmltopdf ";
    
    private static final String MONO_CMD_PATH =  StringUtil.getUploadFilePath() + ConstantUtil.MONO_CMD_PATH;
    
    private static final String DOTNET_CMD_PATH = StringUtil.getUploadFilePath() + ConstantUtil.DOTNET_CMD_PATH;
	
    /**
       * 将office文件转换为html
     * @param inputFileName
     * @param outputFileName
     * @param type
     * @return
     */
	public static boolean officeToHtml(String inputFileName, String outputFileName, String type) {
		//执行转换命令
		boolean result = false;
		try {
			//创建cmd命令
			String cmd = DOTNET_CMD + inputFileName + " " + outputFileName + " " + type;
			logger.debug("officeToHtml cmd is : " + cmd); 
			logger.debug("officeToHtml dir path is : " + DOTNET_CMD_PATH); 
			result = CmdToolUtil.executeCmdToResult(cmd, DOTNET_CMD_PATH, null);
			logger.debug("OfficeUtil.officeToHtml method executed is successful... "); 
		} catch (Exception e) {
			logger.error("OfficeUtil.officeToHtml method executed is error: ", e); 
		}
		return result;
	}
	
	/**
	 * 将word文件转换为html
	 * @param inputFileName /app/file/1.docx
	 * @param outputDirPath /app/file/
	 * @return
	 */
	public static boolean wordToHtml(String inputFileName, String outputDirPath) {
		//执行转换命令
		boolean result = false;
		try {
			//创建cmd命令
			String cmd = MONO_CMD + inputFileName + " " + outputDirPath;
			logger.debug("wordToHtml cmd is : " + cmd); 
			logger.debug("wordToHtml dir path is : " + MONO_CMD_PATH); 
			result = CmdToolUtil.executeCmdToResult(cmd, MONO_CMD_PATH, null);
			logger.debug("OfficeUtil.wordToHtml method executed is successful... "); 
		} catch (Exception e) {
			logger.error("OfficeUtil.wordToHtml method executed is error: ", e); 
		}
		return result;
	}
	
	/**
	 * 将excel文件转成HTML
	 * @param inputFileName /app/file/1.xlsx
	 * @param outputFileName /app/file/1.html
	 * @return
	 */
	public static boolean excelToHtml(String inputFileName, String outputFileName) {
		return officeToHtml(inputFileName, outputFileName, EXCEL_TYPE);
	}
	
	/**
	 * 将PPT文件转成HTML
	 * @param inputFileName /app/file/1.pptx
	 * @param outputFileName /app/file/1.html
	 * @return
	 */
	public static boolean pptToHtml(String inputFileName, String outputFileName) {
		return officeToHtml(inputFileName, outputFileName, PPT_TYPE);
	}
	
	/**
	 * 将HTML转成PDF
	 * @param inputFileName /app/file/1.html
	 * @param outputFileName /app/file/1.pdf
	 * @return
	 */
	public static boolean htmlToPdf(String inputFileName, String outputFileName) {
		//执行转换命令
		boolean result = false;
		try {
			//创建cmd命令
			String cmd = WKHTMLTOPDF_CMD + inputFileName + " " + outputFileName;
			logger.debug("htmlToPdf cmd is : " + cmd); 
			result = CmdToolUtil.executeCmdToResult(cmd, null, null);
			logger.debug("OfficeUtil.htmlToPdf method executed is successful... "); 
		} catch (Exception e) {
			logger.error("OfficeUtil.htmlToPdf method executed is error: ", e); 
		}
		return result;
	}
	
}
