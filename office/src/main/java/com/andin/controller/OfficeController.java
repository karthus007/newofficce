package com.andin.controller;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLDecoder;
import java.util.Base64;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.Part;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.ResponseBody;

import com.andin.model.WaterModel;
import com.andin.service.OfficeService;
import com.andin.thread.OfficeThread;
import com.andin.utils.ConstantUtil;
import com.andin.utils.FileUtil;
import com.andin.utils.StringUtil;
import com.andin.utils.WaterToPdfUtil;

@Controller
@RequestMapping("/office")
public class OfficeController {
	
    private static Logger logger = LoggerFactory.getLogger(OfficeController.class);
    
    private static ExecutorService pool = Executors.newSingleThreadExecutor();
    
	private final static String PDF_DOCX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_DOCX_PATH;
	
	private final static String PDF_XLSX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_XLSX_PATH;
	
	private final static String PDF_PPTX_PATH = StringUtil.getUploadFilePath() + ConstantUtil.PDF_PPTX_PATH;

	@Resource
	private OfficeService officeService;
	
	
	@RequestMapping(value="/pdfToWater", method=RequestMethod.POST)
	@ResponseBody
	public Map<String, Object> pdfToWater(@RequestPart("file") Part part, HttpServletRequest req, HttpServletResponse resp){
		logger.debug("OfficeController.pdfToWater method execute is start...");
		Map<String, Object> map = new HashMap<String, Object>();
		byte[] bytes = "".getBytes();
		try {
			String id = req.getParameter("id") != null ? req.getParameter("id") : "";
			String com = req.getParameter("com") != null ? req.getParameter("com") : "";
			String pass = req.getParameter("pass") != null ? req.getParameter("pass") : "";
			String head = req.getParameter("head") != null ? req.getParameter("head") : "";
			String handler = req.getParameter("handler") != null ? req.getParameter("handler") : "";

			WaterModel water = new WaterModel(handler, head, pass, com, id);
			StringBuffer path = new StringBuffer();
			path.append(StringUtil.getUploadFilePath());
			String fileName = part.getSubmittedFileName();
			int index = fileName.lastIndexOf(".");
			String lastFileName = fileName.substring(0, index);
			String outputFilePath = null;
			String officeInputFilePath = null;
			if(fileName.endsWith(ConstantUtil.PDF)) {
				path.append(ConstantUtil.PDF_PDF_PATH);
				outputFilePath = path.toString() + ConstantUtil.WATER_PATH + fileName;
			}else if(fileName.endsWith(ConstantUtil.DOC) || fileName.endsWith(ConstantUtil.DOCX)) {
				path.append(ConstantUtil.DOCX_PATH);
				officeInputFilePath = PDF_DOCX_PATH + lastFileName + ConstantUtil.PDF;
				outputFilePath = PDF_DOCX_PATH  + ConstantUtil.WATER_PATH + lastFileName + ConstantUtil.PDF;
			}else if(fileName.endsWith(ConstantUtil.XLS) || fileName.endsWith(ConstantUtil.XLSX)) {
				path.append(ConstantUtil.XLSX_PATH);
				officeInputFilePath = PDF_XLSX_PATH + lastFileName + ConstantUtil.PDF;
				outputFilePath = PDF_XLSX_PATH + ConstantUtil.WATER_PATH + lastFileName + ConstantUtil.PDF;
			}else if(fileName.endsWith(ConstantUtil.PPT) || fileName.endsWith(ConstantUtil.PPTX)) {
				path.append(ConstantUtil.PPTX_PATH);
				officeInputFilePath = PDF_PPTX_PATH + lastFileName + ConstantUtil.PDF;
				outputFilePath = PDF_PPTX_PATH + ConstantUtil.WATER_PATH + lastFileName + ConstantUtil.PDF;
			}else {
				map.put(ConstantUtil.RESULT_CODE, ConstantUtil.UPLOAD_FILE_TYPE_ERROR_CODE);
				map.put(ConstantUtil.RESULT_MSG, ConstantUtil.UPLOAD_FILE_TYPE_ERROR_MSG);
				return map;
			}
			String inputFilePath = path.toString() + fileName;
			InputStream in = part.getInputStream();
			byte[] b = new byte[1024*4];
			int len = 0;
			OutputStream os = new FileOutputStream(inputFilePath);
			while ((len = in.read(b)) != -1) {
				os.write(b, 0, len);				
			}
			in.close();
			os.close();
			boolean result = false;
			// 非PDF文件执行PDF任务转换, 生成水印文件
			if(!fileName.endsWith(ConstantUtil.PDF)) {
				Future<Boolean> task = pool.submit(new OfficeThread(fileName));
				if(task.get()) {
					result = WaterToPdfUtil.pdfToWater(officeInputFilePath, outputFilePath, water);
				}
			}else {
				result = WaterToPdfUtil.pdfToWater(inputFilePath, outputFilePath, water);
			}
			if(result) {
				InputStream bin = new FileInputStream(outputFilePath);
				bytes = new byte[bin.available()];
				bin.read(bytes);
				bin.close();
				map.put("data", Base64.getEncoder().encodeToString(bytes));
				map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_SUCCESS_CODE);
				map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_SUCCESS_MSG);				
			}else {
				map.put(ConstantUtil.RESULT_CODE, ConstantUtil.PDF_TO_WATER_ERROR_CODE);
				map.put(ConstantUtil.RESULT_MSG, ConstantUtil.PDF_TO_WATER_ERROR_MSG);
			}
			FileUtil.deleteFilePath(inputFilePath);
			FileUtil.deleteFilePath(outputFilePath);
			logger.debug("OfficeController.pdfToWater method execute is successful...");			
		} catch (Exception e) {
			map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_ERROR_CODE);
			map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_ERROR_MSG);
			logger.error("OfficeController.pdfToWater method execute is error: ", e);
		}
		logger.debug("OfficeController.pdfToWater response is: [resultCode=" + map.get(ConstantUtil.RESULT_CODE) + "],[resultMsg=" + map.get(ConstantUtil.RESULT_MSG) + "]");
		return map;
	}
	
	@RequestMapping(value="/upload", method=RequestMethod.POST)
	@ResponseBody
	public Map<String, Object> upload(HttpServletRequest req, @RequestParam("file") Part part){
		logger.debug("OfficeController.upload method execute is start...");
		Map<String, Object> map = new HashMap<String, Object>();
		try {
			StringBuffer path = new StringBuffer();
			path.append(StringUtil.getUploadFilePath());
			String fileName = part.getSubmittedFileName();
			if(fileName.endsWith(ConstantUtil.DOC) || fileName.endsWith(ConstantUtil.DOCX)) {
				path.append(ConstantUtil.DOCX_PATH);
			}else if(fileName.endsWith(ConstantUtil.XLS) || fileName.endsWith(ConstantUtil.XLSX)) {
				path.append(ConstantUtil.XLSX_PATH);
			}else if(fileName.endsWith(ConstantUtil.PPT) || fileName.endsWith(ConstantUtil.PPTX)) {
				path.append(ConstantUtil.PPTX_PATH);
			}else {
				map.put(ConstantUtil.RESULT_CODE, ConstantUtil.UPLOAD_FILE_TYPE_ERROR_CODE);
				map.put(ConstantUtil.RESULT_MSG, ConstantUtil.UPLOAD_FILE_TYPE_ERROR_MSG);
				return map;
			}
			path.append(fileName);
			InputStream in = part.getInputStream();
			OutputStream os = new FileOutputStream(path.toString());
			byte[] b = new byte[1024*4];
			int len = 0;
			while ((len = in.read(b)) != -1) {
				os.write(b, 0, len);
			}
			in.close();
			os.close();
			map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_SUCCESS_CODE);
			map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_SUCCESS_MSG);
			logger.debug("OfficeController.upload method execute is successful...");
		} catch (Exception e) {
			map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_ERROR_CODE);
			map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_ERROR_MSG);
			logger.error("OfficeController.upload method execute is error: ", e);
		}
		return map;
	}
	
	@RequestMapping(value="/download", method=RequestMethod.GET)
	@ResponseBody
	public Map<String, Object> download(HttpServletRequest req, HttpServletResponse resp){
		logger.debug("OfficeController.download method execute is start...");
		Map<String, Object> map = new HashMap<String, Object>();
		try {
			String fileName = URLDecoder.decode(req.getParameter("name"), "UTF-8");
			String type = "." + URLDecoder.decode(req.getParameter("type"), "UTF-8");
			StringBuffer path = new StringBuffer();
			path.append(StringUtil.getUploadFilePath());
			if(type.equals(ConstantUtil.DOC) || type.equals(ConstantUtil.DOCX)) {
				path.append(ConstantUtil.PDF_DOCX_PATH);
			}else if(type.equals(ConstantUtil.XLS) || type.equals(ConstantUtil.XLSX)) {
				path.append(ConstantUtil.HTML_XLSX_PATH);
			}else if(type.equals(ConstantUtil.PPT) || type.equals(ConstantUtil.PPTX)) {
				path.append(ConstantUtil.PDF_PPTX_PATH);
			}else {
				map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DOWNLOAD_FILE_TYPE_ERROR_CODE);
				map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DOWNLOAD_FILE_TYPE_ERROR_MSG);
				return map;
			}
			path.append(fileName);
	        File file = new File(path.toString());
	        
	        //设置响应头
	        resp.setContentLength((int) file.length());
	        resp.setCharacterEncoding(ConstantUtil.UTF_8);
	        resp.setContentType(ConstantUtil.APPLICATION_OCTET_STREAM);
	        //resp.setHeader("Content-Disposition", "attachment;filename=" + fileName);
	        resp.setHeader("Content-Disposition", "attachment;filename=" + new String( fileName.getBytes(ConstantUtil.UTF_8), "ISO-8859-1"));  
	        BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
	        OutputStream os = resp.getOutputStream();
	        byte[] buff = new byte[1024*4];
	        int len = 0;
	        while ((len = bis.read(buff)) != -1) {
	        	os.write(buff, 0, len);
	        	os.flush();
	        }
	        bis.close();
	        os.close();
	        
			map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_SUCCESS_CODE);
			map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_SUCCESS_MSG);
			logger.debug("OfficeController.download method execute is successful...");
		} catch (Exception e) {
			map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_ERROR_CODE);
			map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_ERROR_MSG);
			logger.error("OfficeController.download method execute is error: ", e.getMessage());
		}
		return map;
	}
	
	@RequestMapping(value="/officeToPdf", method=RequestMethod.GET)
	@ResponseBody
	public Map<String, Object> officeToPdf(@RequestParam("name") String name){
		logger.debug("OfficeController.officeToPdf method execute is start...");
		Map<String, Object> map = new HashMap<String, Object>();
		try {
			Future<Boolean> task = pool.submit(new OfficeThread(name));
			if(task.get()) {
				map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_SUCCESS_CODE);
				map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_SUCCESS_MSG);
			}else {
				map.put(ConstantUtil.RESULT_CODE, ConstantUtil.OFFICE_FILE_CONVERSION_ERROR_CODE);
				map.put(ConstantUtil.RESULT_MSG, ConstantUtil.OFFICE_FILE_CONVERSION_ERROR_MSG);
			}
			logger.debug("OfficeController.officeToPdf method execute is successful...");
		} catch (Exception e) {
			map.put(ConstantUtil.RESULT_CODE, ConstantUtil.DEFAULT_ERROR_CODE);
			map.put(ConstantUtil.RESULT_MSG, ConstantUtil.DEFAULT_ERROR_MSG);
			logger.error("OfficeController.officeToPdf method execute is error: ", e.getMessage());
		}
		return map;
	}
	
	
}
