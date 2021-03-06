package org.easy.excel;


import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.MapUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.easy.excel.config.ExcelDefinition;
import org.easy.excel.config.FieldValue;
import org.easy.excel.exception.ExcelException;
import org.easy.excel.parsing.ExcelExport;
import org.easy.excel.parsing.ExcelHeader;
import org.easy.excel.parsing.ExcelImport;
import org.easy.excel.result.ExcelExportResult;
import org.easy.excel.result.ExcelImportResult;
import org.easy.excel.xml.XMLExcelDefinitionReader;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;

/**
 * Excel上下文支持,只需指定location配置文件路径,即可使用
 * @author lisuo
 *
 */
public class ExcelContext implements ApplicationContextAware{
	
	private ExcelDefinitionReader definitionReader;
	
	/** 用于缓存Excel配置 */
	private Map<String,List<FieldValue>> fieldValueMap = new HashMap<String, List<FieldValue>>();
	
	/**导出*/
	private ExcelExport excelExport;
	/**导入*/
	private ExcelImport excelImport;
	
	/**
	 * @param location 配置文件类路径
	 */
	public ExcelContext(String locations) {
		this(new XMLExcelDefinitionReader(locations));
	}
	
	/**
	 * @param definitionReader 自定义实现ExcelDefinitionReader
	 */
	public ExcelContext (ExcelDefinitionReader definitionReader){
		try {
			if(definitionReader==null){
				throw new ExcelException("definitionReader 不能为空");
			}
			if(MapUtils.isEmpty(definitionReader.getRegistry())){
				throw new ExcelException("definitionReader Registry 不能为空");
			}
			this.definitionReader = definitionReader;
			excelExport = new ExcelExport(definitionReader);
			excelImport = new ExcelImport(definitionReader);
		}catch(ExcelException e){
			throw e;
		}catch (Exception e) {
			throw new ExcelException(e);
		}
		
	}
	
	/**
	 * 创建Excel
	 * @param id 配置ID
	 * @param beans 配置class对应的List
	 * @return Workbook
	 */
	public Workbook createExcel(String id, List<?> beans) {
		return this.createExcel(id, beans, null, null);
	}
	
	/**
	 * 创建Excel部分信息
	 * @param id 配置ID
	 * @param beans 配置class对应的List
	 * @return Workbook
	 */
	public ExcelExportResult createExcelForPart(String id, List<?> beans) {
		return createExcelForPart(id, beans, null, null);
	}
	
	/**
	 * 创建Excel
	 * @param id 配置ID
	 * @param beans 配置class对应的List
	 * @param header 导出之前,在标题前面做出一些额外的操作，比如增加文档描述等,可以为null
	 * @return Workbook
	 */
	public Workbook createExcel(String id, List<?> beans,ExcelHeader header) {
		return createExcel(id, beans, header, null);
	}
	
	/**
	 * 创建Excel部分信息
	 * @param id 配置ID
	 * @param beans 配置class对应的List
	 * @param header 导出之前,在标题前面做出一些额外的操作，比如增加文档描述等,可以为null
	 * @return Workbook
	 */
	public ExcelExportResult createExcelForPart(String id, List<?> beans,ExcelHeader header) {
		return createExcelForPart(id, beans, header, null);
	}
	
	/**
	 * 创建Excel
	 * @param id 配置ID
	 * @param beans 配置class对应的List
	 * @param header 导出之前,在标题前面做出一些额外的操作,比如增加文档描述等,可以为null
	 * @param fields 指定Excel导出的字段(bean对应的字段名称),可以为null
	 * @param workbook 指定excel模板
	 * @return Workbook
	 */
	public Workbook createExcel(String id, List<?> beans,ExcelHeader header,List<String> fields,Workbook workbook) {
		ExcelExportResult result = excelExport.createExcel(id, beans,header,fields,workbook);
		if(result!=null){
			return result.build();
		}
		return null;
	}
	
	/**
	 * 创建Excel
	 * @param id 配置ID
	 * @param beans 配置class对应的List
	 * @param header 导出之前,在标题前面做出一些额外的操作,比如增加文档描述等,可以为null
	 * @param fields 指定Excel导出的字段(bean对应的字段名称),可以为null
	 * @return Workbook
	 */
	public Workbook createExcel(String id, List<?> beans,ExcelHeader header,List<String> fields) {
		return this.createExcel(id, beans, header, fields, null);
	}
	
	/**
	 * 创建Excel部分信息
	 * @param id 配置ID
	 * @param beans 配置class对应的List
	 * @param header 导出之前,在标题前面做出一些额外的操作,比如增加文档描述等,可以为null
	 * @param fields 指定Excel导出的字段(bean对应的字段名称),可以为null
	 * @return Workbook
	 */
	public ExcelExportResult createExcelForPart(String id, List<?> beans,ExcelHeader header,List<String> fields) {
		return this.createExcelForPart(id, beans,header,fields,null);
	}
	
	/**
	 * 创建Excel部分信息
	 * @param id 配置ID
	 * @param beans 配置class对应的List
	 * @param header 导出之前,在标题前面做出一些额外的操作,比如增加文档描述等,可以为null
	 * @param fields 指定Excel导出的字段(bean对应的字段名称),可以为null
	 * @param workbook 指定excel模板
	 * @return Workbook
	 */
	public ExcelExportResult createExcelForPart(String id, List<?> beans,ExcelHeader header,List<String> fields,Workbook workbook) {
		return excelExport.createExcel(id, beans,header,fields,workbook);
	}
	
	/**
	 * 创建Excel,模板信息
	 * @param id	 ExcelXML配置Bean的ID
	 * @param header Excel头信息(在标题之前)
	 * @param fields 指定导出的字段
	 * @return
	 */
	public Workbook createExcelTemplate(String id,ExcelHeader header,List<String> fields,Workbook workbook){
		return excelExport.createExcelTemplate(id, header,fields,workbook);
	}
	
	/**
	 * 创建Excel,模板信息
	 * @param id	 ExcelXML配置Bean的ID
	 * @param header Excel头信息(在标题之前)
	 * @param fields 指定导出的字段
	 * @return
	 */
	public Workbook createExcelTemplate(String id,ExcelHeader header,List<String> fields){
		return this.createExcelTemplate(id, header,fields,null);
	}
	
	/***
	 * 读取Excel信息
	 * @param id 配置ID
	 * @param excelStream Excel文件流
	 * @return ExcelImportResult
	 */
	public ExcelImportResult readExcel(String id, InputStream excelStream) {
		return excelImport.readExcel(id,0, excelStream,null,false);
	}
	
	/***
	 * 读取Excel信息
	 * @param id 配置ID
	 * @param excelStream Excel文件流
	 * @param sheetIndex Sheet索引位
	 * @return ExcelImportResult
	 */
	public ExcelImportResult readExcel(String id, InputStream excelStream,int sheetIndex){
		return excelImport.readExcel(id,0, excelStream,sheetIndex,false);
	}
	
	/***
	 * 读取Excel信息
	 * @param id 配置ID
	 * @param titleIndex 标题索引,从0开始
	 * @param excelStream Excel文件流
	 * @return ExcelImportResult
	 */
	public ExcelImportResult readExcel(String id,int titleIndex, InputStream excelStream) {
		return excelImport.readExcel(id,titleIndex, excelStream,null,false);
	}
	
	/***
	 * 读取Excel信息
	 * @param id 配置ID
	 * @param titleIndex 标题索引,从0开始
	 * @param excelStream Excel文件流
	 * @param multivalidate 是否逐条校验，默认单行出错立即抛出ExcelException，为true时为批量校验,可通过ExcelImportResult.hasErrors,和getErrors获取具体错误信息
	 * @return ExcelImportResult
	 */
	public ExcelImportResult readExcel(String id,int titleIndex, InputStream excelStream,boolean multivalidate) {
		return excelImport.readExcel(id,titleIndex, excelStream,null,multivalidate);
	}
	
	/***
	 * 读取Excel信息
	 * @param id 配置ID
	 * @param titleIndex 标题索引,从0开始
	 * @param excelStream Excel文件流
	 * @param sheetIndex Sheet索引位
	 * @return ExcelImportResult
	 */
	public ExcelImportResult readExcel(String id,int titleIndex, InputStream excelStream,int sheetIndex) {
		return excelImport.readExcel(id,titleIndex, excelStream,sheetIndex,false);
	}
	
	/***
	 * 读取Excel信息
	 * @param id 配置ID
	 * @param titleIndex 标题索引,从0开始
	 * @param excelStream Excel文件流
	 * @param sheetIndex Sheet索引位
	 * @param multivalidate 是否逐条校验，默认单行出错立即抛出ExcelException，为true时为批量校验,可通过ExcelImportResult.hasErrors,和getErrors获取具体错误信息
	 * @return ExcelImportResult
	 */
	public ExcelImportResult readExcel(String id,int titleIndex, InputStream excelStream,int sheetIndex,boolean multivalidate) {
		return excelImport.readExcel(id,titleIndex, excelStream,sheetIndex,multivalidate);
	}
	
	/**
	 * 获取Excel 配置文件中的字段
	 * @param key
	 * @return
	 */
	public List<FieldValue> getFieldValues(String key){
		List<FieldValue> list = fieldValueMap.get(key);
		if(list == null){
			ExcelDefinition def = definitionReader.getRegistry().get(key);
			if(def == null){
				throw new ExcelException("没有找到["+key+"]的配置信息");
			}
			//使用copy方式,避免使用者修改原生的配置信息
			List<FieldValue> fieldValues = def.getFieldValues();
			list = new ArrayList<FieldValue>(fieldValues.size());
			for(FieldValue fieldValue:fieldValues){
				FieldValue val = new FieldValue();
				BeanUtils.copyProperties(fieldValue, val);
				list.add(val);
			}
			fieldValueMap.put(key, list);
		}
		return list;
	}

	@Override
	public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
		this.excelImport.setApplicationContext(applicationContext);
		this.excelExport.setApplicationContext(applicationContext);
	}
	
}
