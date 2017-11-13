package com.jason.util.excel;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.alibaba.dubbo.rpc.RpcContext;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelDownloader {
	
	public static void download(HSSFWorkbook wb, String title) throws IOException{
		HttpServletResponse response = (HttpServletResponse)RpcContext.getContext().getResponse(HttpServletResponse.class);
		HttpServletRequest request = (HttpServletRequest)RpcContext.getContext().getRequest(HttpServletRequest.class);
		BufferedInputStream bins = null;
		BufferedOutputStream bouts = null;
		// 创建文件，返回文件名称
		String fileName = title;
		Date date = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddhhmmss");
		String time = sdf.format(date);
		fileName = new StringBuffer(fileName == null ? "" : fileName).append("_").append(time).append(".xls").toString();
		String path = "E:\\";
		//创建文件夹
		File fileDir = new File(path);
		if (!fileDir.exists()) {
			fileDir.mkdir();
		}
		// 报表写入
		File file = new File(path, fileName);
		if (!file.exists()) {
			file.createNewFile(); 
		}
		//Ie
		if (request.getHeader("User-Agent").toUpperCase().indexOf("MSIE") > 0)
			fileName = URLEncoder.encode(fileName, "UTF-8");
	    else { 
	    	fileName = new String(fileName.getBytes("UTF-8"), "ISO8859-1");
	    }
		FileOutputStream fileOut = null;
		try {
			//将文件给写出流
			fileOut = new FileOutputStream(file);
			//将workbook写入流
			wb.write(fileOut); 
		}catch(Exception e){} finally {
			fileOut.close();
		}
		try {
			response.setContentType("application/x-msdownload;");
			response.setHeader("Content-disposition", "attachment; filename=" + fileName);
			response.setHeader("Content-Length", String.valueOf(file.length()));
			bins = new BufferedInputStream(new FileInputStream(file));
			bouts = new BufferedOutputStream(response.getOutputStream());
		  	byte[] buff = new byte[2048];
		  	int bytesRead;
		  	while (-1 != (bytesRead = bins.read(buff, 0, buff.length))){
		  		bouts.write(buff, 0, bytesRead);
		  	}
	    } catch (Exception e) {
	    	System.out.println("下载文件失败！");
	    	e.printStackTrace();
	    	try
	    	{	
	    		if (bins != null) {
	    			bins.close();
	    		}
	    		if (bouts != null){
	    			bouts.close();
	    		}
	    	}
	    	catch (IOException ioe) {
	    		System.out.println("下载文件关闭流出错！");
	    		e.printStackTrace();
	    	}
	    }
	    finally{
	    	try{
	    		if (bins != null) {
	    			bins.close();
	    		}
	    		if (bouts != null){
	    			bouts.close();
	    		}
	    	}
	    	catch (IOException e) {
	    		System.out.println("下载文件关闭流出错！");
	    		e.printStackTrace();
	    	}
	    	file.delete();
	    }
	}

}