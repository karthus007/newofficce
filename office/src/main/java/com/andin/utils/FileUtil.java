package com.andin.utils;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.tools.zip.ZipEntry;
import org.apache.tools.zip.ZipOutputStream;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 文件操作工具类
 * @author Administrator
 *
 */
public class FileUtil {
	
	private static Logger logger = LoggerFactory.getLogger(FileUtil.class);
	
    private static final String FILE_DEBUG = PropertiesUtil.getProperties("file.debug", null);
	
	/**
	 * 通过不带后缀的文件名、输出的文件路径、需匹配的文件前缀获取ZIP文件
	 * @param fileName 文件名
	 * @param outputDirPath  输出文件路径
	 * @param prefix 匹配前缀
	 * @return
	 */
	public static boolean getFileZipByMatchFileNamePrefix(String fileName, String outputDirPath, String prefix) {
		logger.debug("FileUtil.getFileZipByMatchFileNamePrefix method executed is start, file name is: " + fileName);
		boolean result = false;
		try {
			List<String> fileNameList = getFileNameListByMatchPrefix(prefix, outputDirPath);
			String outputZipNamePath = outputDirPath + fileName + ConstantUtil.ZIP;
			result = getFileZipByFileName(fileNameList, outputZipNamePath, outputDirPath);
			logger.debug("FileUtil.getFileZipByMatchFileNamePrefix method executed is successful... ");
		} catch (Exception e) {
			logger.error("FileUtil.getFileZipByMatchFileNamePrefix method executed is error: ", e);
		}
		return result;
	}
	
	/**
	 * 通过文件名前缀匹配文件夹中的文件名列表
	 * @param prefix 需匹配的前缀
	 * @param dirPath 文件所在路径
	 * @return
	 */
	public static List<String> getFileNameListByMatchPrefix(String prefix, String dirPath){
		logger.debug("FileUtil.getFileNameListByMatch 需匹配的文件名前缀为：" + prefix);
		List<String> list = new ArrayList<String>();
		try {
			File dir = new File(dirPath);
			File[] files = dir.listFiles();
			for (int i = 0; i < files.length; i++) {
				File file = files[i];
				String name = file.getName();
				if(name.startsWith(prefix)) {
					list.add(name);
				}
			}
			logger.debug("FileUtil.getFileNameListByMatch 匹配到的文件名列表为：" + list.toString());
		} catch (Exception e) {
			logger.error("FileUtil.getFileNameListByMatch method executed is error: ", e);
		}
		return list;
	}
	
	
	/**
	 * 将文件夹下的文件名列表打包成zip
	 * @param fileNameList 文件名列表
	 * @param outputZipPath 生成的ZIP文件路径
	 * @param dirPath 需压缩的文件夹路径
	 * @return
	 */
	public static boolean getFileZipByFileName(List<String> fileNameList, String outputZipPath, String dirPath) {
		logger.debug("FileUtil.getFileZipByFileName method executed is start, dir path is: " + dirPath);
		boolean result = false;
		try {
			OutputStream os = new FileOutputStream(outputZipPath);
			ZipOutputStream zos = new ZipOutputStream(new BufferedOutputStream(os));
			zos.setEncoding("GBK");
			logger.debug("FileUtil.getFileZipByFileName method get zip file list is: " + fileNameList.toString());
			for (String name : fileNameList) {
				String fullFileName = dirPath + name;
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
			logger.debug("FileUtil.getFileZipByFileName method executed is successful, output file name path is: " + outputZipPath);
		} catch (Exception e) {
			logger.error("FileUtil.getFileZipByFileName method executed is error: ", e);
		}
		return result;
	}
	
	/**
	 * 文件复制
	 * @param inputFilePath
	 * @param outputFilePath
	 * @return
	 */
	public static boolean copyFilePath(String inputFilePath, String outputFilePath) {
		boolean result = false;
		FileInputStream fis = null;
		FileOutputStream fos = null;
		try{
			fis = new FileInputStream(inputFilePath);
			fos = new FileOutputStream(outputFilePath);
			byte[] bytes = new byte[1024*10];
			int len = 0;
			while((len = fis.read(bytes)) != -1){
				fos.write(bytes, 0, len);
			}
			result = true;
			logger.debug("FileUtil.copyFilePath file copy is successful, path is: " + outputFilePath);
		}catch (Exception e){
			logger.error("FileUtil.copyFilePath method executed is failed: ", e);
		}finally {
			try {
				if(fis != null) {
					fis.close();
				}
				if(fos != null){
					fos.close();
				}
			} catch (Exception e) {
			    logger.error("FileUtil.copyFilePath method close stream is failed: ", e);
			}	
		}
		return result;
	}
	
	/**
	  * 通过文件路径删除文件或文件夹
	 * @param path
	 * @return
	 */
	public static boolean deleteFilePath(String path) {
		if(ConstantUtil.TRUE.equals(FILE_DEBUG)) {
			return true;
		}
		File file = new File(path);
		return deleteFile(file);
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
	
}
