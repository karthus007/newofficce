package com.andin.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class OfficeFileUtil {

	private static Logger logger = LoggerFactory.getLogger(OfficeFileUtil.class);
	
	private final static String DOCX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.DOCX_PATH;
	
	private final static String XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.XLSX_PATH;
	
	private final static String PPTX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PPTX_PATH;

	private final static String PDF_DOCX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_DOCX_PATH;
	
	private final static String PDF_XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_XLSX_PATH;
	
	private final static String PDF_PPTX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_PPTX_PATH;
	
	public static boolean officeToPdf(String inputFileName) {
		boolean result = false;
		try {
			logger.debug("OfficeFileUtil.officeToPdf 转换的文件名为： " + inputFileName);
			int index = inputFileName.lastIndexOf(".");
			String fileName = inputFileName.substring(0, index);
			String fileType = inputFileName.substring(index);
			if(ConstantUtil.DOCX.equals(fileType) || ConstantUtil.DOC.equals(fileType)) {
				//将DOCX文件转换为PDF
				result = OfficeCmdUtil.officeToPdf(DOCX_PATH + inputFileName, PDF_DOCX_PATH + fileName + ConstantUtil.PDF);
				logger.debug("输入文件为：" + inputFileName + ", docx转pdf的结果为：" + result);
				FileUtil.deleteFilePath(DOCX_PATH + inputFileName);
			}else if(ConstantUtil.XLSX.equals(fileType) || ConstantUtil.XLS.equals(fileType)) {
				//将XLSX文件转换为PDF
				result = OfficeCmdUtil.officeToPdf(XLSX_PATH + inputFileName, PDF_XLSX_PATH + fileName + ConstantUtil.PDF);
				logger.debug("输入文件为：" + inputFileName + ", xlsx转pdf的结果为：" + result);
				FileUtil.deleteFilePath(XLSX_PATH + inputFileName);
			}else if(ConstantUtil.PPTX.equals(fileType) || ConstantUtil.PPT.equals(fileType)) {
				//将PPTX文件转换为PDF
				result = OfficeCmdUtil.officeToPdf(PPTX_PATH + inputFileName, PDF_PPTX_PATH + fileName + ConstantUtil.PDF);
				logger.debug("输入文件为：" + inputFileName + ", pptx转pdf的结果为：" + result);
				FileUtil.deleteFilePath(PPTX_PATH + inputFileName);
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
	
}
