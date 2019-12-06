package com.andin.controller;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

import javax.servlet.http.Part;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import com.andin.license.OfficeLicense;
import com.andin.utils.CommonUtil;
import com.andin.utils.ConstantUtil;
import com.andin.utils.StringUtil;

@Controller
@RequestMapping("/license")
public class LicenseController {
		
    private static Logger logger = LoggerFactory.getLogger(LicenseController.class);
	
	@RequestMapping(value="/uploadLicense", method=RequestMethod.POST)
	@ResponseBody
	public Map<String, Object> uploadLicense(@RequestParam("file") Part part){
		logger.debug("TestController.upload method execute is start...");
		Map<String, Object> map = new HashMap<String, Object>();
		try {
			String name = part.getSubmittedFileName();
			if(!ConstantUtil.LICENSE_NAME.equals(name)) {
				map.put(ConstantUtil.RESULT_CODE, ConstantUtil.LICENSE_INVALID_ERROR_CODE);
				map.put(ConstantUtil.RESULT_MSG, ConstantUtil.LICENSE_INVALID_ERROR_MSG);
				return map;
			}
			InputStream in = part.getInputStream();
			byte[] bytes = new byte[in.available()];
			in.read(bytes);
			in.close();
			String license = new String(bytes, ConstantUtil.UTF_8);
			boolean result = OfficeLicense.checkLicense(license);
			if(result) {
				OutputStream los = new FileOutputStream(StringUtil.getUploadFilePath() + ConstantUtil.LICENSE_PATH + ConstantUtil.LICENSE_NAME);
				los.write(bytes);
				los.close();
				CommonUtil.LICENSE_STATUS = true;
				map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_SUCCESS_CODE);
				map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_SUCCESS_MSG);				
			}else {
				map.put(ConstantUtil.RESULT_CODE, ConstantUtil.LICENSE_ERROR_CODE);
				map.put(ConstantUtil.RESULT_MSG, ConstantUtil.LICENSE_ERROR_MSG);		
			}
			logger.debug("TestController.upload method execute is successful...");
		} catch (Exception e) {
			map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_ERROR_CODE);
			map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_ERROR_MSG);
			logger.error("TestController.upload method execute is error: ", e);
		}
		return map;
	}
	
	@RequestMapping(value="/getLicense", method=RequestMethod.GET)
	@ResponseBody
	public Map<String, Object> getLicense(){
		logger.debug("LicenseController.getLicense method execute is start...");
		Map<String, Object> map = new HashMap<String, Object>();
		try {
			if(!CommonUtil.LICENSE_STATUS) {
				map.put("license", StringUtil.getOfficeUUID());				
			}
			map.put("status", CommonUtil.LICENSE_STATUS);
			map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_SUCCESS_CODE);
			map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_SUCCESS_MSG);
			logger.debug("LicenseController.getLicense method execute is successful...");
		} catch (Exception e) {
			map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_ERROR_CODE);
			map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_ERROR_MSG);
			logger.error("LicenseController.getLicense method execute is error: ", e);
		}
		return map;
	}
	
	
}
