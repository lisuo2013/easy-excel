package org.easy.excel.parsing;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.easy.excel.ExcelDefinitionReader;
import org.easy.excel.config.ExcelDefinition;
import org.easy.excel.config.FieldValue;
import org.easy.excel.exception.ExcelException;
import org.easy.excel.result.ExcelExportResult;
import org.springframework.beans.AbstractNestablePropertyAccessor;
import org.springframework.beans.AbstractPropertyAccessor;
import org.springframework.beans.BeanUtils;
import org.springframework.util.TypeUtils;

/**
 * Excel导出实现类
 * @author lisuo
 *
 */
public class ExcelExport extends AbstractExcelResolver{

	
	public ExcelExport(ExcelDefinitionReader definitionReader) {
		super(definitionReader);
	}
	
	//POI 创建cell 样式是有数量限制的，这里为了减少创建，预先创建好需要的单元格样式
	public static class CellStyleHolder{
		
		private CellStyle defaultAlignCellStyle;
		
		private Map<FieldValue,CellStyle> titleCellStyles = new HashMap<>();
		private Map<FieldValue,CellStyle> columnCellStyles = new HashMap<>();
		
		public CellStyleHolder(Workbook workbook,ExcelDefinition excelDefinition) {
			this.init(workbook, excelDefinition);
		}
		
		private void init(Workbook workbook,ExcelDefinition excelDefinition) {
			//文本类型(text)
			DataFormat format = workbook.createDataFormat();
			short textFormat = format.getFormat("@");
			
			if(excelDefinition.getDefaultAlign()!=null) {
				defaultAlignCellStyle = workbook.createCellStyle();
				defaultAlignCellStyle.setAlignment(excelDefinition.getDefaultAlign());
			}
			List<FieldValue> fieldValues = excelDefinition.getFieldValues();
			for (FieldValue fieldValue : fieldValues) {
				if(fieldValue.getAlign()!=null 
						|| fieldValue.getTitleBgColor()!=null 
						|| fieldValue.getTitleFountColor() !=null 
						|| fieldValue.isForceText()){
					CellStyle cellStyle = workbook.createCellStyle();
					if(fieldValue.getAlign()!=null) {
						//设置cell 对齐方式
						cellStyle.setAlignment(fieldValue.getAlign());
					}
					if(fieldValue.getTitleBgColor()!=null) {
						//设置标题背景色
						cellStyle.setFillForegroundColor(fieldValue.getTitleBgColor());
						cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					}
					if(fieldValue.getTitleFountColor()!=null) {
						//设置标题字体色
						Font font = workbook.createFont();
						font.setColor(fieldValue.getTitleFountColor());
						cellStyle.setFont(font);
					}
					if(fieldValue.isForceText()) {
						//设置单元格样式为text
						cellStyle.setDataFormat(textFormat);
					}
					titleCellStyles.put(fieldValue, cellStyle);
				}
				//标题和数据样式不一致,需要个性化的样式
				if(!fieldValue.isUniformStyle()) {
					if(fieldValue.getAlign()!=null 
							||fieldValue.isForceText()) {
						CellStyle colCellStyle = workbook.createCellStyle();
						if(fieldValue.getAlign()!=null) {
							colCellStyle.setAlignment(fieldValue.getAlign());
						}
						if(fieldValue.isForceText()) {
							//设置单元格样式为text
							colCellStyle.setDataFormat(textFormat);
						}
						columnCellStyles.put(fieldValue, colCellStyle);
					}
				}else {
					columnCellStyles.put(fieldValue, titleCellStyles.get(fieldValue));
				}
				
			}
		}
		
		public CellStyle getCellStyle(FieldValue fieldValue,boolean isTitle) {
			if(isTitle) {
				CellStyle cellStyle = titleCellStyles.get(fieldValue);
				return cellStyle != null ? cellStyle:defaultAlignCellStyle;
			}else {
				CellStyle cellStyle = columnCellStyles.get(fieldValue);
				return cellStyle != null ? cellStyle:defaultAlignCellStyle;
			}
		}
		
	}

	/**
	 * 创建导出Excel,如果集合没有数据,返回null
	 * @param id	 ExcelXML配置Bean的ID
	 * @param beans  ExcelXML配置的bean集合
	 * @param header Excel头信息(在标题之前)
	 * @param fields 指定导出的字段
	 * @param workbook
	 * @return
	 * @throws Exception
	 */
	public ExcelExportResult createExcel(String id,List<?> beans,ExcelHeader header,List<String> fields,Workbook workbook) {
		ExcelExportResult exportResult = null;
		if(!CollectionUtils.isEmpty(beans)){
			//从注册信息中获取Bean信息
			ExcelDefinition excelDefinition = definitionReader.getRegistry().get(id);
			if(excelDefinition==null){
				throw new ExcelException("没有找到 ["+id+"] 的配置信息");
			}
			//实际传入的bean类型
			Class<?> realClass = beans.get(0).getClass();
			//传入的类型是excel配置class的类型,或者是它的子类,直接进行生成
			if(realClass==excelDefinition.getClazz() || TypeUtils.isAssignable(excelDefinition.getClazz(),realClass)){
				//导出指定字段的标题不是null,动态创建,Excel定义
				excelDefinition = dynamicCreateExcelDefinition(excelDefinition,fields);
			}
			//传入的类型是excel配置class的类型的父类,那么进行向上转型,只获取配置中父类存在的属性
			else if(TypeUtils.isAssignable(realClass,excelDefinition.getClazz())){
				excelDefinition = extractSuperClassFields(excelDefinition, fields, realClass);
			}else{
				//判断传入的集合与配置文件中的类型拥有共同的父类,如果有则向上转型
				Object superClass = BeanUtil.getEqSuperClass(realClass, excelDefinition.getClazz());
				if(superClass!=Object.class){
					excelDefinition = extractSuperClassFields(excelDefinition, fields, realClass);
				}else{
					throw new ExcelException("传入的参数类型是:"+beans.get(0).getClass().getName()
							+"但是 配置文件的类型是: "+excelDefinition.getClazz().getName()+",参数既不是父类,也不是其相同父类下的子类,无法完成转换");
				}
				
			}
			exportResult = doCreateExcel(excelDefinition,beans,header,workbook);
		}
		return exportResult;
	}
	
	/**
	 * 创建Excel,模板信息
	 * @param id	 ExcelXML配置Bean的ID
	 * @param header Excel头信息(在标题之前)
	 * @param fields 指定导出的字段
	 * @param workbook
	 * @return
	 * @throws Exception
	 */
	public Workbook createExcelTemplate(String id,ExcelHeader header,List<String> fields,Workbook workbook){
		//从注册信息中获取Bean信息
		ExcelDefinition excelDefinition = definitionReader.getRegistry().get(id);
		if(excelDefinition==null){
			throw new ExcelException("没有找到 ["+id+"] 的配置信息");
		}
		excelDefinition = dynamicCreateExcelDefinition(excelDefinition,fields);
		return doCreateExcel(excelDefinition, null, header,workbook).build();
	}
	
	//抽取父类拥用的字段,同时从它的基础只上在进行筛选指定的字段
	private ExcelDefinition extractSuperClassFields(ExcelDefinition excelDefinition,List<String> fields,Class<?> realClass){
		//抽取出父类所拥有的字段
		List<String> fieldNames = BeanUtil.getFieldNames(realClass);
		excelDefinition = dynamicCreateExcelDefinition(excelDefinition, fieldNames);
		//抽取指定的字段
		//导出指定字段的标题不是null,动态创建,Excel定义
		excelDefinition = dynamicCreateExcelDefinition(excelDefinition,fields);
		return excelDefinition;
	}
	
	/**
	 * 动态创建ExcelDefinition
	 * @param excelDefinition 原来的ExcelDefinition
	 * @param fields 
	 * @return
	 */
	private ExcelDefinition dynamicCreateExcelDefinition(ExcelDefinition excelDefinition, List<String> fields) {
		if(!CollectionUtils.isEmpty(fields)){
			ExcelDefinition newDef = new ExcelDefinition();
			BeanUtils.copyProperties(excelDefinition, newDef, "fieldValues");
			List<FieldValue> oldValues = excelDefinition.getFieldValues();
			List<FieldValue> newValues = new ArrayList<FieldValue>(oldValues.size());
			//按照顺序,进行添加
			for(String name:fields){
				for(FieldValue field:oldValues){
					String fieldName = field.getName();
					if(fieldName.equals(name)){
						newValues.add(field);
						break;
					}
				}
			}
			newDef.setFieldValues(newValues);
			return newDef;
		}
		return excelDefinition;
		
	}

	protected ExcelExportResult doCreateExcel(ExcelDefinition excelDefinition, List<?> beans,ExcelHeader header,Workbook workbook){
		// 创建Workbook
		if(workbook==null){
			//XSSFWorkbook支持RichTextString样式
			if(excelDefinition.isRequiredTag()){
				workbook = new XSSFWorkbook();
			}else{
				workbook = new SXSSFWorkbook();
			}
		}
		Sheet sheet = null;
		if(excelDefinition.getSheetname()!=null){
			sheet = workbook.createSheet(excelDefinition.getSheetname());
		}else{
			sheet = workbook.createSheet();
		}
		//创建标题之前,调用buildHeader方法,完成其他数据创建的一些信息
		if(header!=null){
			header.buildHeader(sheet,excelDefinition,beans);
		}
		CellStyleHolder cellStyleHolder = new CellStyleHolder(workbook, excelDefinition);
		Row titleRow = createTitle(excelDefinition,sheet,workbook,cellStyleHolder);
		//如果listBean不为空,创建数据行
		if(beans!=null){
			createRows(excelDefinition, sheet, beans,workbook,titleRow,cellStyleHolder);
		}
		ExcelExportResult exportResult = new ExcelExportResult(excelDefinition, sheet, workbook, titleRow,this,cellStyleHolder);
		return exportResult;
	}

	/**
	 * 创建Excel标题
	 * @param excelDefinition
	 * @param sheet
	 * @return 标题行
	 */
	protected Row createTitle(ExcelDefinition excelDefinition,Sheet sheet,Workbook workbook,CellStyleHolder cellStyleHolder){
		//标题索引号
		int titleIndex = sheet.getPhysicalNumberOfRows();
		Row titleRow = sheet.createRow(titleIndex);
		List<FieldValue> fieldValues = excelDefinition.getFieldValues();
		for(int i=0;i<fieldValues.size();i++){
			FieldValue fieldValue = fieldValues.get(i);
			//设置单元格宽度
			if(fieldValue.getColumnWidth() !=null){
				sheet.setColumnWidth(i, fieldValue.getColumnWidth());
			}
			//如果默认的宽度不为空,使用默认的宽度
			else if(excelDefinition.getDefaultColumnWidth()!=null){
				sheet.setColumnWidth(i, excelDefinition.getDefaultColumnWidth());
			}
			Cell cell = titleRow.createCell(i);
			CellStyle cellStyle = cellStyleHolder.getCellStyle(fieldValue,true);
			if(cellStyle!=null) {
				cell.setCellStyle(cellStyle);
			}
			//处理必填项*色标红
			if(excelDefinition.isRequiredTag() && fieldValue.getTitle().startsWith("*")){
				RichTextString r = new XSSFRichTextString(fieldValue.getTitle());
		        Font ftRed = workbook.createFont();  
		        ftRed.setColor(Font.COLOR_RED);
				r.applyFont(0,1,ftRed);
				cell.setCellValue(r);
			}else{
				setCellValue(cell, fieldValue.getTitle());
			}
		}
		return titleRow;
	}
	
	/**
	 * 创建行
	 * @param excelDefinition
	 * @param sheet
	 * @param beans
	 * @param workbook
	 * @param titleIndex
	 * @throws Exception
	 */
	public void createRows(ExcelDefinition excelDefinition,Sheet sheet,List<?> beans,Workbook workbook,Row titleRow,CellStyleHolder cellStyleHolder){
		int num = sheet.getPhysicalNumberOfRows();
		int startRow = num ;
		for(int i=0;i<beans.size();i++){
			Row row = sheet.createRow(i+num);
			createRow(excelDefinition,row,BeanUtil.buildAccessor(beans.get(i), true),workbook,sheet,titleRow,startRow++,cellStyleHolder);
		}
	}
	
	
	/**
	 * 创建行
	 * @param excelDefinition
	 * @param row
	 * @param bean
	 * @param workbook
	 * @param sheet
	 * @param titleRow
	 * @param rowNum
	 * @throws Exception
	 */
	protected void createRow(ExcelDefinition excelDefinition, Row row, AbstractPropertyAccessor accessor,Workbook workbook,Sheet sheet,Row titleRow,int rowNum,CellStyleHolder cellStyleHolder){
		List<FieldValue> fieldValues = excelDefinition.getFieldValues();
		for(int i=0;i<fieldValues.size();i++){
			FieldValue fieldValue = fieldValues.get(i);
			Object value = BeanUtil.getPropertyValue(accessor,fieldValue.getName(),false);
			//从解析器获取值
			Object instance;
			if(accessor instanceof AbstractNestablePropertyAccessor) {
				instance= ((AbstractNestablePropertyAccessor) accessor).getRootInstance();
			}else {
				instance = ((MapWrapperImpl)accessor).getRootInstance(); 
			}
			Object val = convert(instance,value,fieldValue, Type.EXPORT,rowNum);
			Cell cell = row.createCell(i);
			//cell样式
			CellStyle cellStyle = cellStyleHolder.getCellStyle(fieldValue, false);
			if(cellStyle!=null){
				cell.setCellStyle(cellStyle);
			}
			setCellValue(cell, val);
		}
	}
	
	
}
