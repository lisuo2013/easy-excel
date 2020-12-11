package org.easy.excel.parsing;


import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.easy.excel.ExcelDefinitionReader;
import org.easy.excel.config.ExcelDefinition;
import org.easy.excel.config.FieldValue;
import org.easy.excel.exception.ExcelDataException;
import org.easy.excel.exception.ExcelException;
import org.easy.excel.result.ExcelImportResult;
import org.easy.excel.util.ExcelUtil;
import org.springframework.beans.AbstractPropertyAccessor;
/**
 * Excel导入实现类
 * @author lisuo
 *
 */
public class ExcelImport extends AbstractExcelResolver{
	
	
	public ExcelImport(ExcelDefinitionReader definitionReader) {
		super(definitionReader);
	}
	
	/**
	 * 读取Excel信息
	 * @param id 注册的ID
	 * @param titleIndex 标题索引
	 * @param excelStream Excel文件流
	 * @param sheetIndex Sheet索引位置
	 * @param multivalidate 是否逐条校验，默认单行出错立即抛出ExcelException，为true时为批量校验,可通过ExcelImportResult.hasErrors,和getErrors获取具体错误信息
	 * @return
	 * @throws Exception
	 */
	public ExcelImportResult readExcel(String id, int titleIndex,InputStream excelStream,Integer sheetIndex,boolean multivalidate) {
		//从注册信息中获取Bean信息
		ExcelDefinition excelDefinition = definitionReader.getRegistry().get(id);
		if(excelDefinition==null){
			throw new ExcelException("没有找到 ["+id+"] 的配置信息");
		}
		return doReadExcel(excelDefinition,titleIndex,excelStream,sheetIndex,multivalidate);
	}
	
	protected ExcelImportResult doReadExcel(ExcelDefinition excelDefinition,int titleIndex,InputStream excelStream,Integer sheetIndex,boolean multivalidate) {
		Workbook workbook = ExcelUtil.getWorkBookByStream(excelStream);
		ExcelImportResult result = new ExcelImportResult();
		//读取sheet,sheetIndex参数优先级大于ExcelDefinition配置sheetIndex
		Sheet sheet = ExcelUtil.getSheetAt(workbook, sheetIndex==null?excelDefinition.getSheetIndex():sheetIndex);
		//标题之前的数据处理
		List<List<Object>> header = readHeader(excelDefinition, sheet,titleIndex);
		result.setHeader(header);
		//获取标题
		List<String> titles = readTitle(excelDefinition,sheet,titleIndex);
		//校验标题
		checkTitle(excelDefinition, titles);
		//获取Bean
		List<Object> listBean = readRows(result,excelDefinition,titles, sheet,titleIndex,multivalidate);
		result.setListBean(listBean);
		return result;
	}
	
	/**
	 * 解析标题之前的内容,如果ExcelDefinition中titleIndex 不是0
	 * @param excelDefinition
	 * @param sheet
	 * @return
	 */
	protected List<List<Object>> readHeader(ExcelDefinition excelDefinition,Sheet sheet,int titleIndex){
		List<List<Object>> header = null;
		if(titleIndex!=0){
			header = new ArrayList<List<Object>>(titleIndex);
			for(int i=0;i<titleIndex;i++){
				Row row = sheet.getRow(i);
				if(row == null) {
					continue;
				}
				short cellNum = row.getLastCellNum();
				List<Object> item = new ArrayList<Object>(cellNum);
				for(int j=0;j<cellNum;j++){
					Cell cell = row.getCell(j);
					Object value = getCellValue(cell);
					item.add(value);
				}
				header.add(item);
			}
		}
		return header;
	}
	
	/**
	 * 读取多行
	 * @param result
	 * @param excelDefinition
	 * @param titles
	 * @param sheet
	 * @param titleIndex
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
	protected <T> List<T> readRows(ExcelImportResult result,ExcelDefinition excelDefinition, List<String> titles, Sheet sheet,int titleIndex,boolean multivalidate) {
		//读取数据的总共次数
		int totalNum = sheet.getLastRowNum() - titleIndex;
		List<T> listBean = new ArrayList<T>(totalNum);
		result.setTotalNum(totalNum);
		for (int rowNum = 1; rowNum <= totalNum; rowNum++) {
			try {
				//处理索引位置,为标题索引位置+数据行
				Row row = sheet.getRow(rowNum + titleIndex);
				Object bean = readRow(excelDefinition,row,titles,rowNum);
				listBean.add((T) bean);
			}catch(ExcelDataException e) {
				//应用multivalidate
				if(multivalidate){
					result.getErrors().add(e);
					continue;
				}else{
					throw e;
				}
			}catch(ExcelException e) {
				throw e;
			}
		}
		return listBean;
	}
	
	/**
	 * 读取1行
	 * @param excelDefinition
	 * @param row
	 * @param titles
	 * @param rowNum 第几行
	 * @return
	 * @throws Exception
	 */
	protected Object readRow(ExcelDefinition excelDefinition, Row row, List<String> titles,int rowNum) {
		//创建注册时配置的bean类型
		Object bean = BeanUtil.newInstance(excelDefinition.getClazz());
		AbstractPropertyAccessor accessor = BeanUtil.buildAccessor(bean, true);
		for(FieldValue fieldValue:excelDefinition.getFieldValues()){
			String title = fieldValue.getTitle();
			for (int j = 0; j < titles.size(); j++) {
				//标题或者别名eq
				if(title.equals(titles.get(j)) || fieldValue.getAlias().equals(titles.get(j))){
					Cell cell = row.getCell(j);
					//获取Excel原生value值
					Object value = getCellValue(cell);
					//校验
					validate(fieldValue, value, rowNum,bean);
					if(value != null){
						if(value instanceof String){
							//去除前后空格
							value = value.toString().trim();
						}
						value = super.convert(bean,value, fieldValue, Type.IMPORT,rowNum);
						BeanUtil.setPropertyValue(accessor, fieldValue.getName(), value,false);
					}
					break;
				}
			}
		}
		return bean;
	}

	protected List<String> readTitle(ExcelDefinition excelDefinition, Sheet sheet,int titleIndex) {
		// 获取Excel标题数据
		Row hssfRowTitle = sheet.getRow(titleIndex);
		if(hssfRowTitle==null){
			return null;
		}
		int cellNum = hssfRowTitle.getLastCellNum();
		List<String> titles = new ArrayList<String>(cellNum);
		// 获取标题数据
		for (int i = 0; i < cellNum; i++) {
			Cell cell = hssfRowTitle.getCell(i);
			if(cell == null) {
				continue;
			}
			Object value = getCellValue(cell);
			if(value == null || "".equals(value.toString().trim())) {
				return null;
			}
			titles.add(value.toString());
		}
		return titles;
	}
	
	/**
	 * 数据有效性校验
	 * @param fieldValue
	 * @param value
	 * @param rowNum
	 */
	private void validate(FieldValue fieldValue,Object value,int rowNum,Object refObject){
		if(value == null || (value instanceof String && StringUtils.isBlank(value.toString()))){
			//空校验
			if(!fieldValue.isNull()){
				throw new ExcelDataException("不能为空", rowNum, fieldValue.getAlias(),value,refObject);
			}
		}else{
			if(value instanceof String) {
				//正则校验
				String regex = fieldValue.getRegex();
				if(StringUtils.isNotBlank(regex)){
					String val = value.toString().trim();
					if(!val.matches(regex)){
						String errMsg = fieldValue.getRegexErrMsg()==null?"格式错误":fieldValue.getRegexErrMsg();
						throw new ExcelDataException(errMsg, rowNum, fieldValue.getAlias(),value,refObject);
					}
				}
			}
		}
	}
	
	/**
	 * 校验excel标题与配置的标题是否匹配,如果不包含抛出异常
	 * @param excelDefinition
	 * @param titles
	 */
	private void checkTitle(ExcelDefinition excelDefinition,List<String> titles){
		if(CollectionUtils.isEmpty(titles)){
			throw new ExcelException("标题不能为空");
		}
		List<FieldValue> fieldValues = excelDefinition.getFieldValues();
		//标题校验规则：excel中没有对应的配置标题或别名，同时该属性不能为空，如果为空，允许标题不存在
		for (FieldValue fieldValue : fieldValues) {
			if(!titles.contains(fieldValue.getTitle())){
				if(!titles.contains(fieldValue.getAlias())){
					if(!fieldValue.isNull()){
						throw new ExcelException("标题["+fieldValue.getAlias()+"]在Excel中不存在");
					}
				}
			}
		}
	}
	
}
