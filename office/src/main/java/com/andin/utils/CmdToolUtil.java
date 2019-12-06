package com.andin.utils;

import java.io.BufferedReader;
import java.io.File;
import java.io.InputStream;
import java.io.InputStreamReader;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Administrator
 *
 */
public class CmdToolUtil {
	
    private static Logger logger = LoggerFactory.getLogger(CmdToolUtil.class);

	/**
	 *
	 * @param cmd
	 * @param dirPath
	 * @return
	 * @throws Exception 
	 */
	public static boolean executeCmdToResult(String cmd, String dirPath, String charsetName){
		boolean result = false;
		try {
			Process process = null;
			if (StringUtil.isEmpty(charsetName)) {
				charsetName = "UTF-8";
			}
			if (StringUtil.isEmpty(dirPath)) {
				process = Runtime.getRuntime().exec(cmd);
			} else {
				process = Runtime.getRuntime().exec(cmd, null, new File(dirPath));
			}
			InputStream is = process.getInputStream();
			InputStreamReader in = new InputStreamReader(is);
			BufferedReader reader = new BufferedReader(in);
			@SuppressWarnings("unused")
			String line = "";
			while ((line = reader.readLine()) != null) {
				continue;
			}
			int code = process.waitFor();
			if (code == 0) {
				result = true;
			}
			logger.debug("CmdToolUtil.executeCmdToResult method executed result is: " + result); 
			logger.debug("CmdToolUtil.executeCmdToResult method executed is successful... "); 
		} catch (Exception e) {
			logger.error("CmdToolUtil.executeCmdToResult method executed is error: ", e); 
		}
		return result;
	}
	
	

	/**
	 *
	 * @param cmd
	 * @param dirPath
	 * @param charsetName
	 * @return
	 * @throws Exception
	 */
	public static String executeCmdToContent(String cmd, String dirPath, String charsetName){
		String result = "";
		try {
			Process process = null;
			if (StringUtil.isEmpty(charsetName)) {
				charsetName = "UTF-8";
			}
			if (StringUtil.isEmpty(dirPath)) {
				process = Runtime.getRuntime().exec(cmd);
			} else {
				process = Runtime.getRuntime().exec(cmd, null, new File(dirPath));
			}
			InputStream successis = process.getInputStream();
			InputStreamReader successin = new InputStreamReader(successis);
			BufferedReader successreader = new BufferedReader(successin);
			StringBuilder successcontent = new StringBuilder();
			String successline = "";
			while ((successline = successreader.readLine()) != null) {
				successcontent.append(successline);
			}
			
			InputStream failis = process.getErrorStream();
			InputStreamReader failin = new InputStreamReader(failis);
			BufferedReader failreader = new BufferedReader(failin);
			StringBuilder failcontent = new StringBuilder();
			String failline = "";
			while ((failline = failreader.readLine()) != null) {
				failcontent.append(failline);
			}
			
			process.waitFor();
			if (successcontent.length() != 0) {
				result = successcontent.toString();
			}else{
				result = failcontent.toString();
			}
			successreader.close();
			failreader.close();
			logger.debug("CmdToolUtil.executeCmdToContent method executed result is: " + result); 
			logger.debug("CmdToolUtil.executeCmdToContent method executed is successful... "); 
		} catch (Exception e) {
			logger.error("CmdToolUtil.executeCmdToContent method executed is error: ", e); 
		}
		return result;
	}
	
}
