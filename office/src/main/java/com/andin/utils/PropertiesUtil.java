package com.andin.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Properties;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PropertiesUtil {

	private static final Logger logger = LoggerFactory.getLogger(PropertiesUtil.class);
	
	private static final String LINUX_APP_CONFIG_PATH = "/app/config/";
	
	private static final String WINDOWS_C_APP_CONFIG_PATH = "c:/app/config/";
	
	private static final String WINDOWS_D_APP_CONFIG_PATH = "d:/app/config/";
	
	private static final String CONFIG_FILE_PATH = "config.properties";
	
	private static Properties configPros;
	
	static {
		configPros = getProps(CONFIG_FILE_PATH);
		logger.debug("***PropertiesUtils.init method executed is successful...");
	}
	
	public static Properties getProps(String configpath) {
		Properties props = new Properties();		
		InputStream stream = null;
		try {
			//获取系统的类型
			String systemType = getSystemType();
			if(ConstantUtil.WINDOWS.equals(systemType)) {
				// windows config
				File cfile = new File(WINDOWS_C_APP_CONFIG_PATH + CONFIG_FILE_PATH);
				if(cfile.exists()) {
					stream = new FileInputStream(cfile);
				}else {
					File dfile = new File(WINDOWS_D_APP_CONFIG_PATH + CONFIG_FILE_PATH);
					if(dfile.exists()) {
						stream = new FileInputStream(dfile);
					}else {
						stream = PropertiesUtil.class.getClassLoader().getResourceAsStream(configpath);		
					}
				}
			}else {
				// linux config
				File file = new File(LINUX_APP_CONFIG_PATH + CONFIG_FILE_PATH);
				if(file.exists()) {
					stream = new FileInputStream(file);
				}else {
					stream = PropertiesUtil.class.getClassLoader().getResourceAsStream(configpath);		
				}
			}
			props.load(stream);
			logger.debug("***PropertiesUtils load properties is successful, file name is: " + configpath);
		} catch (Exception e) {
			logger.error("***PropertiesUtils.init method is execute fail: ", e);
		}
		return props;
	}
	
	/**
	 * 根据key从指定文件中配置文件中读取配置,默认获取config文件中的文件
	 * @param key
	 * @param file
	 * @return
	 */
	public static String getProperties(String key, String file) {
		String result = null;
		if(StringUtil.isEmpty(file)) {
			result = configPros.getProperty(key);			
		}else {
			if(ConstantUtil.CONFIG_PROPERTIES.equals(file)) {
				result = configPros.getProperty(key);	
			}
		}
		if(!StringUtil.isEmpty(result)) {
			return result.trim();			
		}else {
			return result;			
		}
	}
	
	/**
	 * 获取系统的类型
	 * @return
	 */
	public static String getSystemType() {
		String type = ConstantUtil.LINUX;
		String name = System.getProperty("os.name").toLowerCase();
		if(name.contains(ConstantUtil.WINDOWS)) {
			type = ConstantUtil.WINDOWS;
		}
		return type;
	}

	public static void main(String[] args) {
		String key = PropertiesUtil.getProperties("te11st", null);
		System.out.println(key);
	}

}
