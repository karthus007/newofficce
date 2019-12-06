package com.andin.utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URI;

import org.apache.http.HttpEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.utils.URIBuilder;
import org.apache.http.entity.ByteArrayEntity;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
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
	
	private static final int HTTP_STATUS_OK = 200;
	
	private static final String CONT = "cont";
	
	private static final String RET = "ret";
	
	private static final int RET_SUCCESS = 1;

	/**
	 * 通过post下载文件
	 * @param mod
	 * @param ac
  	 * @param filePath
	 * @return
	 */
	public static boolean downloadFile(String id, String fileName) {
		boolean result = false;
		CloseableHttpClient client = null;
		CloseableHttpResponse response = null;
		OutputStream os = null;
        try {
        	client = HttpClients.createDefault();
			URI uri = new URIBuilder(OFFICE_HTTP_URI).build();
			HttpPost post = new HttpPost(uri);
			post.addHeader(ConstantUtil.CONTENT_TYPE, ConstantUtil.APPLICATION_JSON_UTF_8);
			String params = "{\"mod\": \"ftranshandle\", \"ac\": \"download\", \"id\": \"" + id + "\"}";
			logger.debug("HttpClientUtil.getDownloadFile method executed params is: " + params);
	        HttpEntity reqEntity = new StringEntity(params);
			post.setEntity(reqEntity);
			response = client.execute(post);
			int status = response.getStatusLine().getStatusCode();
			logger.debug("HttpClientUtil.getDownloadFile method executed response status is: " + status);
			if(status == HTTP_STATUS_OK){
				HttpEntity entity = response.getEntity();
				InputStream in = entity.getContent();
				String filePath = StringUtil.getFilePathByFileName(fileName);
				os = new FileOutputStream(filePath);
				byte[] b = new byte[1024*4];
				int len = 0;
				while((len = in.read(b)) != -1) {
					os.write(b, 0, len);					
				}
			    os.close();
			    result = true;
			    logger.debug("HttpClientUtil.getDownloadFile method executed is successful, file path is: " + filePath);
			}
		} catch (Exception e) {
		    logger.error("HttpClientUtil.getDownloadFile method executed is failed: ", e);
		} finally {
			try {
				if(client != null) {
					client.close();
				}
				if(response!=null){
				    response.close();
				}
			} catch (Exception e) {
			    logger.error("HttpClientUtil.getDownloadFile method close stream is failed: ", e);
			}				
		}
        return result;
	}
	
	/**
	 * 文件上传
	 * @param filePath
	 * @return
	 */
	public static boolean uploadFile(String id, String filePath) {
		boolean result = false;
		CloseableHttpClient client = null;
		CloseableHttpResponse response = null;
        try {
        	//if(filePath.contains(ConstantUtil.HTML_XLSX_PATH)) {
        	//	int index = filePath.lastIndexOf(".");
        	//	filePath = filePath.substring(0, index) + ConstantUtil.ZIP;
        	//}
        	InputStream bis = new FileInputStream(filePath);
        	byte[] arr = new byte[bis.available()];
        	bis.read(arr);
        	bis.close();
        	client = HttpClients.createDefault();
			URI uri = new URIBuilder(OFFICE_HTTP_URI).build();
			HttpPost post = new HttpPost(uri);
			post.addHeader(ConstantUtil.CONTENT_TYPE, ConstantUtil.APPLICATION_OCTET_STREAM);
			post.addHeader(ConstantUtil.HTTP_MOD, ConstantUtil.HTTP_MOD_VALUE);
			post.addHeader(ConstantUtil.HTTP_AC, ConstantUtil.UPLOAD);
			post.addHeader(ConstantUtil.HTTP_ID, id);
			logger.debug("HttpClientUtil.uploadFile method executed params id is: " + id);
	        HttpEntity reqEntity = new ByteArrayEntity(arr);
			post.setEntity(reqEntity);
			response = client.execute(post);
			int status = response.getStatusLine().getStatusCode();
			logger.debug("HttpClientUtil.uploadFile method executed response status is: " + status);
			if(status == HTTP_STATUS_OK){
				HttpEntity respEntity = response.getEntity();
				String resp = EntityUtils.toString(respEntity);
				logger.debug("HttpClientUtil.uploadFile method executed response result is: " + resp);					
				JSONObject json = JSON.parseObject(resp);
				Integer ret = json.getIntValue(RET);
				if(ret == RET_SUCCESS) {
					result = true;
				}
			}
			FileUtil.deleteFilePath(filePath);
		} catch (Exception e) {
		    logger.error("HttpClientUtil.uploadFile method executed is failed: ", e);
		} finally {
			try {
				if(client != null) {
					client.close();
				}
				if(response!=null){
				    response.close();
				}
			} catch (Exception e) {
			    logger.error("HttpClientUtil.uploadFile method close stream is failed: ", e);
			}				
		}
        return result;
	}
	
	/**
	 * 获取待转换的任务
	 * @return
	 */
	public static TaskModel getTask() {
		TaskModel task = null;
		CloseableHttpClient client = null;
		CloseableHttpResponse response = null;
        try {
        	client = HttpClients.createDefault();
			URI uri = new URIBuilder(OFFICE_HTTP_URI).build();
			HttpPost post = new HttpPost(uri);
			post.addHeader(ConstantUtil.CONTENT_TYPE, ConstantUtil.APPLICATION_JSON_UTF_8);
			String params = "{\"mod\": \"ftranshandle\", \"ac\": \"gettask\"}";
			logger.debug("HttpClientUtil.getTaskList method executed params is: " + params);
	        HttpEntity reqEntity = new StringEntity(params);
			post.setEntity(reqEntity);
			response = client.execute(post);
			int status = response.getStatusLine().getStatusCode();
			logger.debug("HttpClientUtil.getTaskList method executed response status is: " + status);
			if(status == HTTP_STATUS_OK){
				HttpEntity respEntity = response.getEntity();
				String resp = EntityUtils.toString(respEntity);
				logger.debug("HttpClientUtil.getTaskList method executed response result is: " + resp);
				JSONObject json = JSON.parseObject(resp);
				Integer ret = json.getIntValue(RET);
				if(ret == RET_SUCCESS) {
					task = json.getObject(CONT, TaskModel.class);
				}
			}
		} catch (Exception e) {
		    logger.error("HttpClientUtil.getTaskList method executed is failed: ", e);
		} finally {
			try {
				if(client != null) {
					client.close();
				}
				if(response!=null){
				    response.close();
				}
			} catch (Exception e) {
			    logger.error("HttpClientUtil.getTaskList method close stream is failed: ", e);
			}				
		}
        return task;
	}
	
	/**
	 * 修改任务的状态
	 * @param id
	 * @param stat
	 * @return
	 */
	public static boolean updateTaskStatus(String id) {
		boolean result = false;
		CloseableHttpClient client = null;
		CloseableHttpResponse response = null;
        try {
        	client = HttpClients.createDefault();
			URI uri = new URIBuilder(OFFICE_HTTP_URI).build();
			HttpPost post = new HttpPost(uri);
			post.addHeader(ConstantUtil.CONTENT_TYPE, ConstantUtil.APPLICATION_JSON_UTF_8);
			String params = "{\"mod\": \"ftranshandle\", \"ac\": \"uptaskstat\", \"id\": \"" + id + "\", \"stat\": 5}";
			logger.debug("HttpClientUtil.updateTaskStatus method executed params is: " + params);
	        HttpEntity reqEntity = new StringEntity(params);
			post.setEntity(reqEntity);
			response = client.execute(post);
			int status = response.getStatusLine().getStatusCode();
			logger.debug("HttpClientUtil.updateTaskStatus method executed response status is: " + status);
			if(status == 200){
				HttpEntity respEntity = response.getEntity();
				String resp = EntityUtils.toString(respEntity);
				logger.debug("HttpClientUtil.updateTaskStatus method executed response result is: " + resp);
				JSONObject json = JSON.parseObject(resp);
				Integer ret = json.getIntValue(RET);
				if(ret == RET_SUCCESS) {
					result = true;
				}
			}
		} catch (Exception e) {
		    logger.error("HttpClientUtil.updateTaskStatus method executed is failed: ", e);
		} finally {
			try {
				if(client != null) {
					client.close();
				}
				if(response!=null){
				    response.close();
				}
			} catch (Exception e) {
			    logger.error("HttpClientUtil.updateTaskStatus method close stream is failed: ", e);
			}				
		}
        return result;
	}
	
	
	public static void main(String[] args) throws Exception {
		System.out.println("===开始获取任务===");
		//getTask();
		//downloadFile("5d68c1a07eaa3cd57e8b4cb3", "5d68c1a07eaa3cad398b4e0b.doc");
		//updateTaskStatus("5d68c1a07eaa3cd57e8b4cb3");
		//uploadFile("5d68c0367eaa3cd16d8b4a63", "d:/app/ccc.txt");
		System.out.println("===结束获取任务===");
	}
	
}
