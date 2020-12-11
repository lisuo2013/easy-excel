package org.easy.excel.parsing;


import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.easy.excel.ExcelDefinitionReader;
import org.easy.excel.config.FieldValue;
import org.easy.excel.exception.ExcelException;
import org.easy.excel.util.ExcelUtil;
import org.springframework.context.ApplicationContext;

/**
 * Excel抽象解析器
 * 
 * @author lisuo
 *
 */
public abstract class AbstractExcelResolver implements CellValueConverter{

	protected ExcelDefinitionReader definitionReader;
	
	protected ApplicationContext ctx;

	/** 注册字段解析信息 */
	private Map<String,CellValueConverter> cellValueConverters = new HashMap<String, CellValueConverter>();
	private DefaultCellValueConverter defaultCellValueConverter = new DefaultCellValueConverter();
	
	public AbstractExcelResolver(ExcelDefinitionReader definitionReader) {
		this.definitionReader = definitionReader;
	}


	/**
	 * 设置Cell单元的值
	 * 
	 * @param cell
	 * @param value
	 */
	protected void setCellValue(Cell cell, Object value) {
		ExcelUtil.setCellValue(cell, value);
	}

	/**
	 * 获取cell值
	 * 
	 * @param cell
	 * @return
	 */
	protected Object getCellValue(Cell cell) {
		return ExcelUtil.getCellValue(cell);
	}
	
	//默认实现
	@Override
	public Object convert(Object bean,Object value, FieldValue fieldValue, Type type,int rowNum){
		if(value !=null){
			//解析器实现，读取数据
			String convName = fieldValue.getCellValueConverterName();
			if(convName==null){
				return defaultCellValueConverter.convert(bean, value, fieldValue, type, rowNum);
			}else{
				//自定义
				CellValueConverter conv = cellValueConverters.get(convName);
				if(conv == null){
					synchronized(this){
						if(conv == null){
							conv = getCellValueConverter(convName);
							cellValueConverters.put(convName, conv);
						}
					}
					conv = cellValueConverters.get(convName);
				}
				Object val = conv.convert(bean,value, fieldValue, type, rowNum);
				if(val != null){
					return val;
				}
			}
		}
		return fieldValue.getDefaultValue();

	}
	
	/**
	 * 注入spring context
	 * @param applicationContext
	 */
	public void setApplicationContext(ApplicationContext applicationContext) {
		this.ctx = applicationContext;
	}
	
	/**
	 * 如果是spring环境从容器中获取转换器
	 * @param convName
	 * @return
	 */
	protected CellValueConverter getCellValueConverter(String convName){
		try{
			CellValueConverter bean = null;
			if(ctx!=null){
				bean = (CellValueConverter) ctx.getBean(Class.forName(convName));
			}else{
				bean =  (CellValueConverter) BeanUtil.newInstance(Class.forName(convName));
			}
			return bean;
		}catch(Exception e){
			throw new ExcelException(e);
		}
	}
	
}
