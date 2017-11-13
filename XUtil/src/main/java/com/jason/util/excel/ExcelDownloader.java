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

/**
 * Excel下载器
 * @author Jason
 */
public class ExcelDownloader {
	
	//TODO 实际开发环境中用LOG代替System.out
	//static LogHandler LOG = LogHandler.getLogHandler(Download.class);

	/**
	 * 浏览器下载
	 * 注意接口采用GET请求 / 参数通过URL传递 / 前端使用window.location.href = url
	 * @param title
	 * @throws IOException
	 */
	public static void browserDownload(HSSFWorkbook wb) throws IOException{
		HttpServletResponse response = (HttpServletResponse)RpcContext.getContext().getResponse(HttpServletResponse.class);
		HttpServletRequest request = (HttpServletRequest)RpcContext.getContext().getRequest(HttpServletRequest.class);
		BufferedInputStream bins = null;
		BufferedOutputStream bouts = null;
		//在D盘创建临时文件，将文件放入流后删除文件
		Date date = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddhhmmss");
		String time = sdf.format(date);
		String fileName = "Excel";
		fileName = new StringBuffer(fileName == null ? "" : fileName).append("_").append(time).append(".xls").toString();
		String path = "D:\\";
		File fileDir = new File(path);
		if (!fileDir.exists()) {
			fileDir.mkdir();
		}
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
			//将文件、workbook写入流
			fileOut = new FileOutputStream(file);
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

	/**
	 * 本地指定路径下载
	 * @date 2017年11月13日 下午3:21:08
	 * @author Jason
	 * @param filePath
	 * @param wb
	 */
	public static void localDownload(String filePath , HSSFWorkbook wb){
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(filePath);
			wb.write(out);
		} catch (Exception e) {
			System.out.println("文件写入出错");
			e.printStackTrace();
		} finally{
			if(out != null){
				try {
					out.close();
				} catch (IOException e) {
					System.out.println("下载文件关闭流出错");
					e.printStackTrace();
				}
			}
		}
	}
	
}