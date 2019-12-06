package com.andin.utils;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 文件操作工具类
 * @author Administrator
 *
 */
public class FileUtil {
	
	private static Logger logger = LoggerFactory.getLogger(FileUtil.class);

	private final static String HTML_XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.HTML_XLSX_PATH;
		
	/**
	 * 获取包含文件名的不带后缀的文件名列表
	 * @param fileName
	 * @return
	 */
	private static List<String> getXlsxHtmlFileNameList(String fileName) throws Exception{
		logger.debug("FileUtil.getXlsxHtmlFileNameList 需匹配的文件名为：" + fileName);
		List<String> list = new ArrayList<String>();
		//获取html/xlsx/文件夹下的文件名列表
		File dir = new File(HTML_XLSX_PATH);
		File[] files = dir.listFiles();
		for (int i = 0; i < files.length; i++) {
			File file = files[i];
			String name = file.getName();
			if(name.startsWith(fileName)) {
				list.add(name);
			}
		}
		logger.debug("FileUtil.getXlsxHtmlFileNameList 匹配到的文件名列表为：" + list.toString());
		return list;
	}
	
	/**
	 * 将通过不在后缀匹配到的文件名都压缩成一个zip包
	 * @param fileName
	 * @param zipFileName
	 * @return
	 */
	public static boolean getHtmlFileZipByFileName(String fileName) {
		logger.debug("FileUtil.getHtmlFileZipByFileName method executed is start, file name is: " + fileName);
		boolean result = false;
		try {
			List<String> list = getXlsxHtmlFileNameList(fileName);
			logger.debug("FileUtil.getHtmlFileZipByFileName method get zip file list is: " + list.toString());
			OutputStream os = new FileOutputStream(HTML_XLSX_PATH + fileName + ConstantUtil.ZIP);
			ZipOutputStream zos = new ZipOutputStream(new BufferedOutputStream(os));
			for (String name : list) {
				String fullFileName = HTML_XLSX_PATH + name;
				File file = new File(fullFileName);
				if(file.isDirectory()) {
					File[] files = file.listFiles();
					for (File item : files) {
					    InputStream in = new FileInputStream(item);
					    ZipEntry entry = new ZipEntry(file.getName() + "/" + item.getName());
					    zos.putNextEntry(entry);
					    byte[] bytes = new byte[1024*1024];
					    int len = 0;
					    while ((len = in.read(bytes)) != -1) {
					    	zos.write(bytes, 0, len);
					    }
					    in.close();
					    zos.closeEntry();
					}
				}else {
				    InputStream in = new FileInputStream(file);
				    ZipEntry entry = new ZipEntry(file.getName());
				    zos.putNextEntry(entry);
				    byte[] bytes = new byte[1024*1024];
				    int len = 0;
				    while ((len = in.read(bytes)) != -1) {
				    	zos.write(bytes, 0, len);
				    }
				    in.close();
				    zos.closeEntry();
				}
				deleteFile(file);
			}
			zos.close();
			result = true;
			logger.debug("FileUtil.getHtmlFileZipByFileName method executed is successful... ");
		} catch (Exception e) {
			logger.error("FileUtil.getHtmlFileZipByFileName method executed is error: ", e);
		}
		return result;
	}
	
	
	/**
	  * 通过文件删除文件或文件夹
	 * @param file
	 * @return
	 */
	public static boolean deleteFile(File file) {
		boolean result = false;
		try {
			if(file.exists()) {
				if(file.isDirectory()) {
					File[] list = file.listFiles();
					for (int i = 0; i < list.length; i++) {
						deleteFile(list[i]);
					}	
				}
				file.delete();
				logger.debug("FileUtil.deleteFile file delete is successful, path is: " + file.getAbsolutePath());
			}else {
				logger.debug("FileUtil.deleteFile file is not exist, path is: " + file.getAbsolutePath());
			}
			result = true;
		} catch (Exception e) {
			result = false;
			logger.error("FileUtil.deleteFile method executed is error: ", e);
		}
		return result;
	}
	
	/**
	  * 通过文件路径删除文件或文件夹
	 * @param path
	 * @return
	 */
	public static boolean deleteFilePath(String path) {
		File file = new File(path);
		return deleteFile(file);
	}
	
}
