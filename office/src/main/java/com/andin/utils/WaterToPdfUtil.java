package com.andin.utils;

import java.io.FileOutputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.andin.model.WaterModel;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGState;
import com.itextpdf.text.pdf.PdfLayer;
import com.itextpdf.text.pdf.PdfName;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfWriter;

public class WaterToPdfUtil {

    private static Logger logger = LoggerFactory.getLogger(WaterToPdfUtil.class);
	
	public static final String WATER_FONT_PATH = StringUtil.getUploadFilePath() + ConstantUtil.WATER_FONT_PATH;
	
	public static final String WATER_IMAGE_PATH = StringUtil.getUploadFilePath() + ConstantUtil.WATER_IMAGE_PATH;
	
	public static final String WATERMARK = "Watermark";
	/* --- 生成的PDF是否为纵向，默认纵向 ---  */
	public static final boolean IS_LEVEL_PDF = false;
	/* --- 水印的文件颜色 ---  */
	public static final BaseColor WATER_COLOR = BaseColor.GRAY;
	
    /**
          * 给PDF文件添加水印
     * @param inputFilePath /app/file/1.pdf
     * @param outputFilePath /app/file/2.pdf
     * @param water 水印信息
     * @return
     */
    public static boolean pdfToWater(String inputFilePath, String outputFilePath, WaterModel water){
		logger.debug("WaterToPdfUtil.pdfToWater method is start, water is: " + water.toString()); 
    	boolean result = false;
    	try{
    	    String WATER_COM = water.getCom();
    	    String WATER_PASS = water.getPass();
    	    String SIZE_TOP_TEXT = "合同编号:" + water.getId();
    	    String SIZE_BOTTOM_TEXT = "承办人:" + water.getHandler() + "  " + "部门负责人:" + water.getHead();
        	PdfReader pdfReader = new PdfReader(inputFilePath);
        	int numberOfPages = pdfReader.getNumberOfPages();
        	FileOutputStream outputStream = new FileOutputStream(outputFilePath, true);
        	PdfStamper pdfStamper = new PdfStamper(pdfReader, outputStream);
        	PdfContentByte waterMarkContent;
            BaseFont bf = BaseFont.createFont(WATER_FONT_PATH, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            if(!StringUtil.isEmpty(WATER_PASS)) {
            	byte[] bytes = WATER_PASS.getBytes();
            	pdfStamper.setEncryption(bytes, bytes, PdfWriter.ALLOW_MODIFY_CONTENTS, false);
            }            
            PdfLayer layer = new PdfLayer(WATERMARK, pdfStamper.getWriter());
            for (int i = 1; i <= numberOfPages; i++){
                waterMarkContent = pdfStamper.getOverContent(i);
                Rectangle rect = pdfReader.getPageSize(i);
                waterMarkContent.beginLayer(layer);
                PdfName pdfname = new PdfName(WATERMARK);
                waterMarkContent.saveState();
                waterMarkContent.beginMarkedContentSequence(pdfname);
                int nLen = WATER_COM.length();
                float nFontsize = 0;
                if(nLen <= 10){
                    nFontsize = 20;
                    nLen = 5 * nLen;
                }else if (nLen > 10 && nLen <= 15){
                    nFontsize = 18;
                    nLen = 4 * nLen;
                }else if (nLen > 15 && nLen <= 20){
                    nFontsize = 16;
                    nLen = 4 * nLen;
                }else{
                    nFontsize = 12;
                    nLen = 3 * nLen;
                }
                //字体路径需要传入
                waterMarkContent.setFontAndSize(bf, nFontsize);
                PdfGState gState = new PdfGState();
                gState.setFillOpacity(0.4f);
                waterMarkContent.setGState(gState);
                waterMarkContent.setColorFill(WATER_COLOR);
                waterMarkContent.beginText();
                Image img = Image.getInstance(WATER_IMAGE_PATH);
                float ftWidth, ftHeight;
                if (rect.getWidth()> rect.getHeight()){
                    ftWidth = rect.getHeight();
                    ftHeight = rect.getWidth();
                }else{
                    ftWidth = rect.getWidth();
                    ftHeight = rect.getHeight();
                }
                img.scalePercent(100f);
 	            img.scaleAbsolute(50, 50);
 	            img.setAbsolutePosition(50,80);
 	            waterMarkContent.addImage(img);
 	            img.setAbsolutePosition(50,ftHeight - 330);
 	            waterMarkContent.addImage(img);
 	            img.setAbsolutePosition(ftWidth/2 + 50,80);
 	            waterMarkContent.addImage(img);
 	            img.setAbsolutePosition(ftWidth/2 + 50,ftHeight - 330);
 	            waterMarkContent.addImage(img);
 	            waterMarkContent.showTextAligned(PdfContentByte.ALIGN_LEFT, WATER_COM, 100, 130, 52);
 	            waterMarkContent.showTextAligned(PdfContentByte.ALIGN_LEFT, WATER_COM, ftWidth/2 + 100, 130, 52);
 	            waterMarkContent.showTextAligned(PdfContentByte.ALIGN_LEFT, WATER_COM, 100, ftHeight - 280, 52);
 	            waterMarkContent.showTextAligned(PdfContentByte.ALIGN_LEFT, WATER_COM, ftWidth/2 + 100, ftHeight - 280, 52);
 	            if (nFontsize-4>12){
                    nFontsize = 12;
                }
                //生成的PDF是否为纵向，默认为纵向
                if (IS_LEVEL_PDF){
                    if (rect.getWidth() > rect.getHeight()){
                        waterMarkContent.showTextAligned(PdfContentByte.ALIGN_RIGHT, SIZE_BOTTOM_TEXT, rect.getHeight() - 30, 10, 0);
                        waterMarkContent.showTextAligned(PdfContentByte.ALIGN_RIGHT, SIZE_TOP_TEXT, rect.getHeight() - 30, rect.getWidth() - 20, 0);
                    }else{
                        waterMarkContent.showTextAligned(PdfContentByte.ALIGN_RIGHT, SIZE_BOTTOM_TEXT, rect.getWidth() - 30, 10, 0);
                        waterMarkContent.showTextAligned(PdfContentByte.ALIGN_RIGHT, SIZE_TOP_TEXT, rect.getWidth() - 30, rect.getHeight() - 20, 0);
                    }
                }else{
                    waterMarkContent.showTextAligned(PdfContentByte.ALIGN_RIGHT, SIZE_BOTTOM_TEXT, rect.getWidth() - 30, 10, 0);
                    waterMarkContent.showTextAligned(PdfContentByte.ALIGN_RIGHT, SIZE_TOP_TEXT, rect.getWidth() - 30, rect.getHeight() - 20, 0);
                }

                waterMarkContent.endText();
                waterMarkContent.endLayer();
            }
            pdfStamper.close();
            pdfReader.close();
            outputStream.close();
            result = true;
			logger.debug("WaterToPdfUtil.pdfToWater method executed is successful... "); 
        }catch (Exception e){
			logger.error("WaterToPdfUtil.pdfToWater method executed is error: ", e); 
        }
        return result;
    }
    
}