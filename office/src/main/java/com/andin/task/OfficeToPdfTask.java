package com.andin.task;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import com.andin.model.TaskModel;
import com.andin.utils.CommonUtil;
import com.andin.utils.HttpClientUtil;
import com.andin.utils.OfficeFileUtil;
import com.andin.utils.PropertiesUtil;
import com.andin.utils.StringUtil;

/**
 * OFFICE转PDF文件的定时任务
 * @author Administrator
 *
 */
@Component
@EnableScheduling
public class OfficeToPdfTask {
	
	private static final Integer TASK_THREAD_COUNT = Integer.valueOf(PropertiesUtil.getProperties("task.thread.count", null));
    
	private static Logger logger = LoggerFactory.getLogger(OfficeToPdfTask.class);

	private static ExecutorService pool = Executors.newFixedThreadPool(TASK_THREAD_COUNT);
	
	public static volatile Map<String, Integer> map = new ConcurrentHashMap<>();

	@Scheduled(cron = "*/5 * * * * ?")/** 每五秒触发一次 **/
	public void getOfficeTaskListToPdf() throws Exception{
		if(CommonUtil.LICENSE_STATUS) {
			logger.debug("OfficeToPdfTask.getOfficeTaskListToPdf method executed is start...");
			int tcount = ((ThreadPoolExecutor) pool).getActiveCount();
			if(tcount < TASK_THREAD_COUNT) {
				TaskThread thread = new TaskThread();
				pool.execute(thread);	
			}
		}else {
			logger.debug("OfficeToPdfTask.getOfficeTaskListToPdf license is authorization failed...");
		}
	}
	
	/**
	  * 启动线程转换任务
	 * @author Administrator
	 *
	 */
	public class TaskThread extends Thread{
		
		@Override 
		public void run() {
			//获取任务列表
			TaskModel task = HttpClientUtil.getTask();
			if(task != null) {
				long startTime = System.currentTimeMillis();
				Boolean downloadResult = false;
				Boolean officeToPdfResult = false;
				Boolean uploadResult = false;
				Boolean updateResult = false;
				logger.debug("OfficeToPdfTask.getOfficeTaskListToPdf method task params is: " + task.toString());
				String taskId = task.getId();
				String name = task.getFilename();
				String fileType = StringUtil.getFileTypeByType(task.getFiletype());
				String fileName = name + fileType;
				//通过文件ID从PHP下载文件
				downloadResult = HttpClientUtil.downloadFile(taskId, fileName);
				if(downloadResult) {
					//开始OFFICE转换PDF, excel转html
					officeToPdfResult = OfficeFileUtil.officeToPdf(fileName);
					if(officeToPdfResult) {
						//通过文件名获取转换好的PDF文件的路径, EXCEL为html路径
						String filePath = StringUtil.getPdfFilePathByFileName(fileName, name);
						//上传文件到PHP
						uploadResult = HttpClientUtil.uploadFile(taskId, filePath);
						if(uploadResult) {
							//更新任务的转换状态						
							updateResult = HttpClientUtil.updateTaskStatus(taskId, 5);
						}
					}else {
						Integer taskIdCount = OfficeToPdfTask.map.get(taskId);
						if(taskIdCount == null) {
							taskIdCount = 0;
						}
						if(taskIdCount < 3) {
							try {
								Thread.sleep(120000);
							} catch (Exception e) {
								logger.debug("OfficeToPdfTask.getOfficeTaskListToPdf Thread sleep is error, ", e);
							}
							//更新任务的转换状态						
							updateResult = HttpClientUtil.updateTaskStatus(taskId, 0);
							taskIdCount += 1;
							OfficeToPdfTask.map.put(taskId, taskIdCount);
						}else {
							OfficeToPdfTask.map.remove(taskId);
						}
					}
				}
				String result = "[downloadResult=" + downloadResult + "], [officeToPdfResult=" + officeToPdfResult + "], [uploadResult=" + uploadResult + "], [updateResult=" + updateResult + "]";
				logger.debug("OfficeToPdfTask.getOfficeTaskListToPdf method executed task result is: " + result);
				long endTime = System.currentTimeMillis();
			    logger.debug("OfficeToPdfTask.getOfficeTaskListToPdf method executed spend time is: " + (endTime - startTime)/1000 + "s");
			    
			}
		}
		
	}
	
}
