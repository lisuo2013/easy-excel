package org.easy.excel.util;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

/**
 * Excel 下载工具类,提供原生的Servlet下载环境
 * @author lisuo
 *
 */
public class ExcelDownLoadUtil {

	/** 文件后缀 */
	public static final String FILE_SUFFIX = ".xlsx";

	/** 文件编码 */
	public static final String UTF8 = "UTF-8";
	/** 用户浏览器关键字：IE */
	private static final String USER_AGENT_IE [] = {"MSIE", "Trident", "Edge"};

	private static final String CONTENT_TYPE = "application/x-excel";
	
	/**
	 * 下载Excel,解决中文乱码问题,如果Workbook为空，执行alert(emptyMessage);
	 * @param workbook POI Workbook
	 * @param excelName Excel名字（不需要后缀，支持中文处理）
	 * @param emptyMessage workbook为空提示的信息
	 * @param request
	 * @param response
	 * @throws IOException
	 */
	public static void downLoadExcel(Workbook workbook,String excelName,String emptyMessage,HttpServletRequest request,HttpServletResponse response)throws IOException{
		if (workbook != null) {
			String excelFileName = encodeDownloadFileName(request, excelName + FILE_SUFFIX);
			response.setContentType(CONTENT_TYPE);
			response.setHeader("Content-Disposition", "attachment; filename=\"" + excelFileName + "\";target=_blank");
			OutputStream out = response.getOutputStream();
			workbook.write(out);
			workbook.close();
			out.flush();
			out.close();
		} else {
			response.setContentType("text/html; charset=utf-8");
			PrintWriter writer = response.getWriter();
			writer.print("<script language='javascript'>alert('"+emptyMessage+"');</script>");
			writer.flush();
			writer.close();
		}
	}
	
	/**
	 * 下载Excel（原生文件下载,只是把响应头按照excel设置,解决中文乱码问题）
	 * @param ins 原生的excel文件流
	 * @param excelName 文件名称
	 * @param request
	 * @param response
	 * @throws IOException
	 */
	public static void downLoadExcel(InputStream ins,String excelName,HttpServletRequest request,HttpServletResponse response)throws IOException{
		String excelFileName = encodeDownloadFileName(request, excelName);
		response.setContentType(CONTENT_TYPE);
		response.setHeader("Content-Disposition", "attachment; filename=\"" + excelFileName + "\";target=_blank");
		OutputStream out = response.getOutputStream();
		byte[] bs = IOUtils.toByteArray(ins);
		ins.close();
		out.write(bs);
		out.flush();
		out.close();
	}
	
	/**
	 * 根据不同的浏览器设置下载文件名称的编码
	 * @param request
	 * @param fileName
	 * @return 文件名称
	 */
	public static String encodeDownloadFileName(HttpServletRequest request, String fileName) {
		String userAgent = request.getHeader("User-Agent");
		boolean isIe = false;
		for (String ie : USER_AGENT_IE) {
			if(userAgent.indexOf(ie) > 0){
				isIe = true;
				break;
			}
		}
		if (isIe) {// 用户在用IE
			try {
				return URLEncoder.encode(fileName, UTF8);
			} catch (UnsupportedEncodingException ignore) {}
		} else {
			try {
				return new String(fileName.getBytes(UTF8), "ISO-8859-1");
			} catch (UnsupportedEncodingException ignore) {
			}
		}
		return fileName;
	}
	
}
