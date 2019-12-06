package com.andin.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * office文件转pdf工具类
 * @author Administrator
 *
 */
public class OfficeCmdUtil {
	
	private static Logger logger = LoggerFactory.getLogger(OfficeCmdUtil.class);
	
    private static final String OFFICE_CMD = "unoconv -f pdf -o ";
	
    /**
          * 将office文件转换为pdf
     * @param inputFileName
     * @param outputFileName
     * @param type
     * @return
     */
	public static boolean officeToPdf(String inputFileName, String outputFileName) {
		//执行转换命令
		boolean result = false;
		try {
			//创建cmd命令
			String cmd = OFFICE_CMD + outputFileName + " " + inputFileName;
			logger.debug("OfficeCmdUtil.officeToPdf cmd is : " + cmd); 
			result = CmdToolUtil.executeCmdToResult(cmd, null, null);
			logger.debug("OfficeCmdUtil.officeToPdf method executed is successful... "); 
		} catch (Exception e) {
			logger.error("OfficeCmdUtil.officeToPdf method executed is error: ", e); 
		}
		return result;
	}
	
}
