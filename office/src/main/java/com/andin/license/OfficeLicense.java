package com.andin.license;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Date;

import javax.crypto.Cipher;
import javax.crypto.KeyGenerator;
import javax.crypto.spec.SecretKeySpec;

import org.apache.commons.codec.binary.Base64;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSONObject;
import com.andin.utils.ConstantUtil;
import com.andin.utils.StringUtil;

public class OfficeLicense {
	
    private static Logger logger = LoggerFactory.getLogger(OfficeLicense.class);
	
    private static final String ALGORITHMSTR = "AES/ECB/PKCS5Padding";  
    
    /**
          * 将bytes数组进行base64加密
     * @param bytes
     * @return
     */
    public static String base64Encode(byte[] bytes){
        Base64 base64 = new Base64();
        return base64.encodeToString(bytes);
    }
    
    /**
          * 将加密字符串进行base64解密
     * @param base64Code
     * @return
     * @throws Exception
     */
    public static byte[] base64Decode(String base64Code) throws Exception{
        Base64 base64 = new Base64();
        return base64.decode(base64Code);
    }  
    
    /**
     * AES加密
     * @param content 待加密字符串
     * @param encryptKey 加密秘钥
     * @return
     * @throws Exception
     */
    public static String aesEncrypt(String content, String encryptKey) throws Exception {  
        return base64Encode(aesEncryptToBytes(content, encryptKey));
    }
    
    /**
     * AES解密
     * @param encryptContent 待解密字符串
     * @param decryptKey 解密秘钥
     * @return
     * @throws Exception
     */
    public static String aesDecrypt(String encryptContent, String decryptKey) throws Exception {  
        return aesDecryptByBytes(base64Decode(encryptContent), decryptKey);
    }
    
    /**
     * AES将字符串加密成byte数组
     * @param content
     * @param encryptKey
     * @return
     * @throws Exception
     */
    private static byte[] aesEncryptToBytes(String content, String encryptKey) throws Exception {  
        KeyGenerator kgen = KeyGenerator.getInstance("AES");  
        kgen.init(128);  
        Cipher cipher = Cipher.getInstance(ALGORITHMSTR);  
        cipher.init(Cipher.ENCRYPT_MODE, new SecretKeySpec(encryptKey.getBytes(), "AES"));  
        return cipher.doFinal(content.getBytes("utf-8"));  
    }
    
    /**
     * AES将byte数组解密成字符串
     * @param encryptBytes
     * @param decryptKey
     * @return
     * @throws Exception
     */
    private static String aesDecryptByBytes(byte[] encryptBytes, String decryptKey) throws Exception {  
        KeyGenerator kgen = KeyGenerator.getInstance("AES");  
        kgen.init(128);  
        Cipher cipher = Cipher.getInstance(ALGORITHMSTR);  
        cipher.init(Cipher.DECRYPT_MODE, new SecretKeySpec(decryptKey.getBytes(), "AES"));  
        byte[] decryptBytes = cipher.doFinal(encryptBytes);  
        return new String(decryptBytes);  
    }
    
    /**
          * 通过key生成lincense
     * @param key
     * @return
     * @throws Exception
     */
    private static void getLicenseFileByKey(String key, String licenseFilePath) throws Exception{
    	String license = null;
    	JSONObject json = new JSONObject();
    	json.put(ConstantUtil.LICENSE_KEY, key);
    	json.put(ConstantUtil.COM_KEY, ConstantUtil.COM_VALUE);
    	json.put(ConstantUtil.CREATE_TIME, new Date().getTime());
    	license = key + aesEncrypt(json.toJSONString(), key);
    	OutputStream os = new FileOutputStream(licenseFilePath + ConstantUtil.LICENSE_NAME);
    	os.write(license.getBytes());
    	os.close();
    }
    
    /**
          * 通过key检测license是否成功
     * @param license
     * @param key
     * @return
     */
    public static boolean checkLicense(String license) {
    	boolean result = false;
    	try {
			String data = aesDecrypt(license.substring(32), license.substring(0, 32));
	        JSONObject info = JSONObject.parseObject(data);
	        String uuid = info.getString(ConstantUtil.LICENSE_KEY);
	        result = StringUtil.checkOfficeLicenseByUUID(uuid);
    	} catch (Exception e) {
    		logger.error("***OfficeLicense.checkLicense method executed is error: ", e);
		}
    	return result;
    }
	
    
    public static void main(String[] args) throws Exception {
    	String key = "0736063dc126739482102446f3235d43";
    	String path = "d:/app/";
        getLicenseFileByKey(key, path);
        System.out.println("====get license method executed is successful====");        
    }
}