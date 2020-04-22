package com.andin.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class OfficeFileUtil {

	private static Logger logger = LoggerFactory.getLogger(OfficeFileUtil.class);
	
    private static final String OFFICE_EXCEL_TYPE = PropertiesUtil.getProperties("office.excel.type", null);
    
    private static final String OFFICE_CONVERT_TYPE = PropertiesUtil.getProperties("office.convert.type", null);
	
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
					if(ConstantUtil.ASPOSE_CONVERT_TYPE.equals(OFFICE_CONVERT_TYPE)) {
						result = OfficeUtil.asposeWordToPdf(input, output);
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
				if(ConstantUtil.OFFICE_EXCEL_TO_HTML.equals(OFFICE_EXCEL_TYPE)) {
					// html文件路径
					String output = StringUtil.getOutputFilePathByFileName(fileName + ConstantUtil.HTML);
					for (int i = 0; i < repeatCount; i++) {
						//将EXCEL文件转换为HTML
						result = OfficeUtil.asposeExcelToHtml(input, output);
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.asposeExcelToHtml 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
					}
					logger.debug("输入文件为：" + inputFileName + ", xlsx转html的结果为：" + result);
					if(result) {
						//获取生成的文件名前缀
						String prefix = fileName;
						//将html文件压缩成zip包
						result = FileUtil.getFileZipByMatchFileNamePrefix(fileName, HTML_XLSX_PATH, prefix);
						logger.debug("输入文件为：" + inputFileName + ", html文件压缩成zip的结果为：" + result);
					}
				// excel to pdf
				}else if(ConstantUtil.OFFICE_EXCEL_TO_PDF.equals(OFFICE_EXCEL_TYPE)) {
					String output = StringUtil.getOutputFilePathByFileName(inputFileName);
					//将PPTX文件转换为PDF
					for (int i = 0; i < repeatCount; i++) {
						//存在输出文件则删除
						FileUtil.deleteFilePath(output);
						//将EXCEL文件转换为PDF
						
						// aspose convert to pdf
						if(ConstantUtil.ASPOSE_CONVERT_TYPE.equals(OFFICE_CONVERT_TYPE)) {
							result = OfficeUtil.asposeExcelToPdf(input, output);
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
					//将EXCEL文件转换为PNG
					String output = StringUtil.getOutputFilePathByFileName(fileName + ConstantUtil.PNG);
					for (int i = 0; i < repeatCount; i++) {
						//将EXCEL文件转换为PNG
						result = OfficeUtil.asposeExcelToImage(input, output);
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.asposeExcelToImage 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + XLSX_PATH + inputFileName);
					}
					logger.debug("输入文件为：" + inputFileName + ", xlsx转png的结果为：" + result);
					FileUtil.deleteFilePath(input);
					if(result) {
						//获取生成的文件名前缀
						String prefix = fileName + "-";
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
					// aspose convert to pdf
					if(ConstantUtil.ASPOSE_CONVERT_TYPE.equals(OFFICE_CONVERT_TYPE)) {
						result = OfficeUtil.asposePptxToPdf(input, output);
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
					if(ConstantUtil.ASPOSE_CONVERT_TYPE.equals(OFFICE_CONVERT_TYPE)) {
						result = OfficeUtil.asposeWordToPdf(input, output);
					// default office convert to pdf
					}else {
						result = OfficeUtil.officeWordToPdf(input, output);
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
				if(ConstantUtil.OFFICE_EXCEL_TO_HTML.equals(OFFICE_EXCEL_TYPE)) {
					// html文件路径
					String output = StringUtil.getOutputFilePathByFileName(fileName + ConstantUtil.HTML);
					for (int i = 0; i < repeatCount; i++) {
						//将EXCEL文件转换为HTML
						result = OfficeUtil.asposeExcelToHtml(input, output);
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.asposeExcelToHtml 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
					}
					logger.debug("输入文件为：" + inputFileName + ", xlsx转html的结果为：" + result);
					if(result) {
						//获取生成的文件名前缀
						String prefix = fileName;
						//将html文件压缩成zip包
						result = FileUtil.getFileZipByMatchFileNamePrefix(fileName, HTML_XLSX_PATH, prefix);
						logger.debug("输入文件为：" + inputFileName + ", html文件压缩成zip的结果为：" + result);
					}
				// excel to pdf
				}else if(ConstantUtil.OFFICE_EXCEL_TO_PDF.equals(OFFICE_EXCEL_TYPE)) {
					String output = StringUtil.getOutputFilePathByFileName(inputFileName);
					//将PPTX文件转换为PDF
					for (int i = 0; i < repeatCount; i++) {
						//存在输出文件则删除
						FileUtil.deleteFilePath(output);
						//将EXCEL文件转换为PDF
						result = OfficeUtil.asposeExcelToPdf(input, output);
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.windowsOfficeToPdf 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + input);
					}
				// default excel to png
				}else {
					String output = StringUtil.getOutputFilePathByFileName(fileName + ConstantUtil.PNG);
					for (int i = 0; i < repeatCount; i++) {
						//将EXCEL文件转换为PNG
						result = OfficeUtil.asposeExcelToImage(input, output);
						if(result) {
							break;
						}
						logger.debug("OfficeFileUtil.asposeExcelToImage 正在第" + (i + 1) + "次重试转换..., 文件名称为：" + XLSX_PATH + inputFileName);
					}
					logger.debug("输入文件为：" + inputFileName + ", xlsx转png的结果为：" + result);
					FileUtil.deleteFilePath(input);
					if(result) {
						//获取生成的文件名前缀
						String prefix = fileName + "-";
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
					result = OfficeUtil.asposePptxToPdf(input, output);
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

}
