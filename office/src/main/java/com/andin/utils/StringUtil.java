package com.andin.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.net.Inet4Address;
import java.net.InetAddress;
import java.net.NetworkInterface;
import java.util.Enumeration;
import java.util.Random;
import java.util.UUID;

import com.andin.license.OfficeLicense;

public class StringUtil {
	
    private static final String FILE_PATH = PropertiesUtil.getProperties("file.path", null);
    
    private static final String FILE_STATUS = PropertiesUtil.getProperties("file.status", null);
	
	private final static String UPLOAD_PATH = "/upload/";
	
    private static final String RANDOM_INFO = "abcdefghijklmnopqrstuvwxyz0123456789";
	
    
	/**
	  * 根据长度生成指定的字符串
	 * @param length
	 * @return
	 */
	public static String getRandomString(int length){
	    Random random = new Random();
	    StringBuffer sb = new StringBuffer();
	    for(int i=0; i<length; i++){
	      int number = random.nextInt(36);
	      sb.append(RANDOM_INFO.charAt(number));
	    }
	    return sb.toString();
	}
    
	/**
	  * 判断字符串是否为空
	 * @param name
	 * @return
	 */
	public static boolean isEmpty(String name){
		boolean result = false;
		if(name == null || name == "" || name == "null" || name.trim().length() == 0) {
			result = true;
		}
		return result;
	}
	
	/**
	 * 获取项目的根路径
	 * @return
	 */
	public static String getUploadFilePath(){
		if(ConstantUtil.TRUE.equals(FILE_STATUS)) {
			return FILE_PATH + UPLOAD_PATH;
		}else {
			return System.getProperty("user.dir").replace("\\", "/") + UPLOAD_PATH;
		}
	}
	
	/**
	 * 通过文件名获取文件的存储路径
	 * @param fileName
	 * @return
	 */
	public static String getFilePathByFileName(String fileName) {
		StringBuffer path = new StringBuffer();
		path.append(getUploadFilePath());
		if(fileName.endsWith(ConstantUtil.DOC) || fileName.endsWith(ConstantUtil.DOCX)) {
			path.append(ConstantUtil.DOCX_PATH);
		}else if(fileName.endsWith(ConstantUtil.XLS) || fileName.endsWith(ConstantUtil.XLSX)) {
			path.append(ConstantUtil.XLSX_PATH);
		}else if(fileName.endsWith(ConstantUtil.PPT) || fileName.endsWith(ConstantUtil.PPTX)) {
			path.append(ConstantUtil.PPTX_PATH);
		}else if(fileName.endsWith(ConstantUtil.PDF)) {
			path.append(ConstantUtil.PDF_PDF_PATH);
		}
		path.append(fileName);
		return path.toString();
	}
	
	/**
	 * 通过文件名获取转换好后PDF的文件路径
	 * @param fileName
	 * @return
	 */
	public static String getPdfFilePathByFileName(String fileName, String name) {
		StringBuffer path = new StringBuffer();
		path.append(getUploadFilePath());
		if(fileName.endsWith(ConstantUtil.DOC) || fileName.endsWith(ConstantUtil.DOCX)) {
			path.append(ConstantUtil.PDF_DOCX_PATH);
		}else if(fileName.endsWith(ConstantUtil.XLS) || fileName.endsWith(ConstantUtil.XLSX)) {
			path.append(ConstantUtil.PDF_XLSX_PATH);
		}else if(fileName.endsWith(ConstantUtil.PPT) || fileName.endsWith(ConstantUtil.PPTX)) {
			path.append(ConstantUtil.PDF_PPTX_PATH);
		}else if(fileName.endsWith(ConstantUtil.PDF)) {
			path.append(ConstantUtil.PDF_PDF_PATH);
		}
		path.append(name);
		path.append(ConstantUtil.PDF);
		return path.toString();
	}
	
	/**
	 * 通过类型获取文件名的后缀
	 * @param type
	 * @return
	 */
	public static String getFileTypeByType(Integer type) {
		String fileType = "";
		switch (type) {
		case 82:
			fileType = ConstantUtil.PPT;
			break;
		case 83:
			fileType = ConstantUtil.DOC;
			break;
		case 84:
			fileType = ConstantUtil.PDF;
			break;
		case 85:
			fileType = ConstantUtil.XLS;
			break;
		default:
			fileType = ConstantUtil.PDF;
			break;
		}
		return fileType;
	}
	
	/**
	 * 通过IP生成UUID
	 * @return
	 * @throws Exception
	 */
	public static String getOfficeUUID() throws Exception {
		StringBuffer officeid = new StringBuffer(32);
		String ip = getInet4Address();
		String uuid = UUID.randomUUID().toString().replace("-", "");
		String[] arr = ip.split("\\.");
		for (int i = 0; i < arr.length; i++) {
			String rule = String.valueOf((Integer.valueOf(arr[i]) + 111));
			officeid.append(uuid.substring(i*8, i*8 + 2));
			officeid.append(rule.charAt(0));
			officeid.append(uuid.substring(i*8 + 2, i*8 + 3));
			officeid.append(rule.charAt(1));
			officeid.append(uuid.substring(i*8 + 3, i*8 + 4));
			officeid.append(rule.charAt(2));
			officeid.append(uuid.substring(i*8 + 4, i*8 + 5));
		}
		return officeid.toString();
	}
	
	public static boolean checkOfficeLicenseByUUID(String uuid) {
		boolean result = false;
		try {
			String host = getInet4Address();
			//通过uuid获取ip
			StringBuffer ip = new StringBuffer(12);
			ip.append(uuid.charAt(2));
			ip.append(uuid.charAt(4));
			ip.append(uuid.charAt(6));
			ip.append(uuid.charAt(10));
			ip.append(uuid.charAt(12));
			ip.append(uuid.charAt(14));
			ip.append(uuid.charAt(18));
			ip.append(uuid.charAt(20));
			ip.append(uuid.charAt(22));
			ip.append(uuid.charAt(26));
			ip.append(uuid.charAt(28));
			ip.append(uuid.charAt(30));
			String[] arr = host.split("\\.");
			StringBuffer officeip = new StringBuffer(12);
			for (int i = 0; i < arr.length; i++) {
				String rule = String.valueOf((Integer.valueOf(arr[i]) + 111));
				officeip.append(rule);
			}
			if(officeip.toString().equals(ip.toString())) {
				 result = true;
			}
		} catch (Exception e) {
			 result = false;
		}
		return result;
	}
	
	/**
	 * 获取项目的license的状态
	 * @return
	 */
	public static boolean getLicenseStatus() {
		boolean licenseStatus = false;
		try {
			File file = new File(StringUtil.getUploadFilePath() + ConstantUtil.LICENSE_PATH + ConstantUtil.LICENSE_NAME);
			if(file.exists()) {
				InputStream in = new FileInputStream(file);
				byte[] bytes = new byte[in.available()];
				in.read(bytes);
				in.close();
				String license = new String(bytes, ConstantUtil.UTF_8);
				licenseStatus = OfficeLicense.checkLicense(license);
			}
		} catch (Exception e) {
			licenseStatus = false;
		}
		return licenseStatus;
	}
	
	/**
	  * 获取IP地址
	 * @return
	 * @throws Exception
	 */
	public static String getInet4Address() throws Exception{
		String ip = null;
		Enumeration<NetworkInterface> nis = NetworkInterface.getNetworkInterfaces();
		for (; nis.hasMoreElements();) {
			NetworkInterface ni = nis.nextElement();
			Enumeration<InetAddress> ias = ni.getInetAddresses();
			for (;ias.hasMoreElements();) {
				InetAddress ia = ias.nextElement();
				if (ia instanceof Inet4Address && !ia.getHostAddress().equals("127.0.0.1")) {
					ip = ia.getHostAddress();
				}
			}
		}
		return ip;
	}

}
