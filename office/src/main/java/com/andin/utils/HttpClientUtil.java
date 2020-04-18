package com.andin.utils;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.andin.model.TaskModel;

/**
 * HTTP调用第三方服务获取数据
 * @author Administrator
 *
 */
public class HttpClientUtil {
	
    private static Logger logger = LoggerFactory.getLogger(HttpClientUtil.class);
    
	private static final String OFFICE_HTTP_URI = PropertiesUtil.getProperties("office.uri", null);
	
	private static final String OFFICE_TYPE = PropertiesUtil.getProperties("office.type", null);
	
	private static final int HTTP_STATUS_OK = 200;
	
	private static final String HTTP_POST = "POST";
	
	private static final String CONT = "cont";
	
	private static final String RET = "ret";
	
	private static final int RET_SUCCESS = 1;

	
	/**
	 * 获取待转换的任务
	 * @return
	 */
	public static TaskModel getTask() {
		long startTime = System.currentTimeMillis();
		TaskModel task = null;
        try {
        	URL url = new URL(OFFICE_HTTP_URI);
        	HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        	connection.setRequestMethod(HTTP_POST);
        	connection.setConnectTimeout(10000);
        	connection.setDoOutput(true);
            connection.setDoInput(true);
            connection.setUseCaches(false);
            connection.setInstanceFollowRedirects(true);
            // 设置请求头
            connection.setRequestProperty(ConstantUtil.CONTENT_TYPE, ConstantUtil.APPLICATION_JSON_UTF_8);
            connection.connect();
            // 设置请求参数
			String params = "{\"mod\": \"ftranshandle\", \"ac\": \"gettask\", \"type\": \"" + OFFICE_TYPE + "\"}";
			logger.debug("HttpClientUtil.getTask method executed params is: " + params);
            OutputStream out = connection.getOutputStream();
            out.write(params.getBytes());
            out.flush();
            out.close();

            int status = connection.getResponseCode();
            logger.debug("HttpClientUtil.getTaskList method executed response status is: " + status);
            if(status == HTTP_STATUS_OK){
                BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
                String line = "";
                String resp = "";
                while ((line = reader.readLine()) != null) {
                	resp += line;
                }
                reader.close();
                logger.debug("HttpClientUtil.getTaskList method executed response result is: " + resp);
                JSONObject json = JSON.parseObject(resp);
				Integer ret = json.getIntValue(RET);
				if(ret == RET_SUCCESS) {
					task = json.getObject(CONT, TaskModel.class);
				}
            }
            // 断开连接
            connection.disconnect();
		} catch (Exception e) {
		    logger.error("HttpClientUtil.getTaskList method executed is failed: ", e);
		}
        long endTime = System.currentTimeMillis();
	    logger.debug("HttpClientUtil.getTaskList method executed spend time is: " + (endTime - startTime)/1000 + "s");

        return task;
	}
	
	/**
	 * 通过post下载文件
	 * @param mod
	 * @param ac
  	 * @param filePath
	 * @return
	 */
	public static boolean downloadFile(String id, String fileName) {
		long startTime = System.currentTimeMillis();
		boolean result = false;
        try {
        	URL url = new URL(OFFICE_HTTP_URI);
        	HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        	connection.setRequestMethod(HTTP_POST);
        	connection.setConnectTimeout(10000);
        	connection.setDoOutput(true);
            connection.setDoInput(true);
            connection.setUseCaches(false);
            connection.setInstanceFollowRedirects(true);
            // 设置请求头
            connection.setRequestProperty(ConstantUtil.CONTENT_TYPE, ConstantUtil.APPLICATION_JSON_UTF_8);
            connection.connect();
            // 设置请求参数
			String params = "{\"mod\": \"ftranshandle\", \"ac\": \"download\", \"id\": \"" + id + "\"}";
			logger.debug("HttpClientUtil.getDownloadFile method executed params is: " + params);
            OutputStream out = connection.getOutputStream();
            out.write(params.getBytes());
            out.flush();
            out.close();

            int status = connection.getResponseCode();
            logger.debug("HttpClientUtil.getDownloadFile method executed response status is: " + status);
            if(status == HTTP_STATUS_OK){
            	InputStream in = connection.getInputStream();
				String filePath = StringUtil.getFilePathByFileName(fileName);
				OutputStream os = new FileOutputStream(filePath);
				byte[] b = new byte[1024*4];
				int len = 0;
				while((len = in.read(b)) != -1) {
					os.write(b, 0, len);					
				}
			    os.close();
			    in.close();
			    result = true;
			    logger.debug("HttpClientUtil.getDownloadFile method executed is successful, file path is: " + filePath);
            }
            // 断开连接
            connection.disconnect();
		} catch (Exception e) {
		    logger.error("HttpClientUtil.getDownloadFile method executed is failed: ", e);
		}
        long endTime = System.currentTimeMillis();
	    logger.debug("HttpClientUtil.getDownloadFile method executed spend time is: " + (endTime - startTime)/1000 + "s");
        return result;
	}
	
	/**
	 * 文件上传
	 * @param filePath
	 * @return
	 */
	public static boolean uploadFile(String id, String filePath) {
		long startTime = System.currentTimeMillis();
		boolean result = false;
        try {
        	if(filePath.contains(ConstantUtil.PDF_XLSX_PATH)) {
        		int index = filePath.lastIndexOf(".");
        		filePath = filePath.substring(0, index) + ConstantUtil.ZIP;
        	}
        	InputStream bis = new FileInputStream(filePath);
        	byte[] arr = new byte[bis.available()];
        	bis.read(arr);
        	bis.close();

        	URL url = new URL(OFFICE_HTTP_URI);
        	HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        	connection.setRequestMethod(HTTP_POST);
        	connection.setConnectTimeout(10000);
        	connection.setDoOutput(true);
            connection.setDoInput(true);
            connection.setUseCaches(false);
            connection.setInstanceFollowRedirects(true);
            // 设置请求头
            connection.setRequestProperty(ConstantUtil.CONTENT_TYPE, ConstantUtil.APPLICATION_OCTET_STREAM);
            connection.setRequestProperty(ConstantUtil.HTTP_MOD, ConstantUtil.HTTP_MOD_VALUE);
            connection.setRequestProperty(ConstantUtil.HTTP_AC, ConstantUtil.UPLOAD);
            connection.setRequestProperty(ConstantUtil.HTTP_ID, id);
            connection.connect();
			logger.debug("HttpClientUtil.uploadFile method executed params id is: " + id);
            OutputStream out = connection.getOutputStream();
            out.write(arr);
            out.flush();
            out.close();

            int status = connection.getResponseCode();
            logger.debug("HttpClientUtil.uploadFile method executed response status is: " + status);
            if(status == HTTP_STATUS_OK){
                BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
                String line = "";
                String resp = "";
                while ((line = reader.readLine()) != null) {
                	resp += line;
                }
                reader.close();
                logger.debug("HttpClientUtil.uploadFile method executed response result is: " + resp);
                JSONObject json = JSON.parseObject(resp);
				Integer ret = json.getIntValue(RET);
				if(ret == RET_SUCCESS) {
					result = true;
				}
            }
            // 断开连接
            connection.disconnect();
			FileUtil.deleteFilePath(filePath);
		} catch (Exception e) {
		    logger.error("HttpClientUtil.uploadFile method executed is failed: ", e);
		}
        
        long endTime = System.currentTimeMillis();
	    logger.debug("HttpClientUtil.uploadFile method executed spend time is: " + (endTime - startTime)/1000 + "s");
        return result;
	}
	
	/**
	 * 修改任务的状态
	 * @param id
	 * @param stat
	 * @return
	 */
	public static boolean updateTaskStatus(String id, Integer stat) {
		long startTime = System.currentTimeMillis();
		boolean result = false;
        try { 
        	URL url = new URL(OFFICE_HTTP_URI);
        	HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        	connection.setRequestMethod(HTTP_POST);
        	connection.setConnectTimeout(10000);
        	connection.setDoOutput(true);
            connection.setDoInput(true);
            connection.setUseCaches(false);
            connection.setInstanceFollowRedirects(true);
            // 设置请求头
            connection.setRequestProperty(ConstantUtil.CONTENT_TYPE, ConstantUtil.APPLICATION_JSON_UTF_8);
            connection.connect();
            // 设置请求参数
			String params = "{\"mod\": \"ftranshandle\", \"ac\": \"uptaskstat\", \"id\": \"" + id + "\", \"stat\": " + stat + "}";
			logger.debug("HttpClientUtil.updateTaskStatus method executed params is: " + params);
            OutputStream out = connection.getOutputStream();
            out.write(params.getBytes());
            out.flush();
            out.close();

            int status = connection.getResponseCode();
            logger.debug("HttpClientUtil.updateTaskStatus method executed response status is: " + status);
            if(status == HTTP_STATUS_OK){
                BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
                String line = "";
                String resp = "";
                while ((line = reader.readLine()) != null) {
                	resp += line;
                }
                reader.close();
                logger.debug("HttpClientUtil.updateTaskStatus method executed response result is: " + resp);
                JSONObject json = JSON.parseObject(resp);
				Integer ret = json.getIntValue(RET);
				if(ret == RET_SUCCESS) {
					result = true;
				}
            }
            // 断开连接
            connection.disconnect();
		} catch (Exception e) {
		    logger.error("HttpClientUtil.updateTaskStatus method executed is failed: ", e);
		}
        long endTime = System.currentTimeMillis();
	    logger.debug("HttpClientUtil.updateTaskStatus method executed spend time is: " + (endTime - startTime)/1000 + "s");
        return result;
	}
	
	
	public static void main(String[] args) throws Exception {
		System.out.println("===开始获取任务===");
		getTask();
		//downloadFile("5d68c1a07eaa3cd57e8b4cb3", "5d68c1a07eaa3cad398b4e0b.doc");
		updateTaskStatus("5d68c1a07eaa3cd57e8b4cb3", 5);
		//uploadFile("5d68c0367eaa3cd16d8b4a63", "d:/app/ccc.txt");
		System.out.println("===结束获取任务===");
	}
	
}
