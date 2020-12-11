package org.easy.excel.parsing;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.easy.excel.config.FieldValue;
import org.easy.excel.exception.ExcelDataException;

/**
 * 默认的CellValueConverter转换器实现
 * @author lisuo
 *
 */
public class DefaultCellValueConverter implements CellValueConverter{

	@Override
	public Object convert(Object bean, Object value, FieldValue fieldValue, Type type, int rowNum){
		//执行默认
		String name = fieldValue.getName();
		String pattern = fieldValue.getPattern();
		String format = fieldValue.getFormat();
		DecimalFormat decimalFormat = fieldValue.getDecimalFormat();
		if (StringUtils.isNotBlank(pattern)) {
			String [] patterns = StringUtils.split(pattern, ",");
			if (Type.EXPORT == type) {
				//导出使用第一个pattern
				return DateFormatUtils.format((Date) value, patterns[0]);
			} else if (Type.IMPORT == type) {
				if (value instanceof String) {
					Date date = parseDate((String) value, patterns);
					if(date==null){
						StringBuilder errMsg = new StringBuilder("[");
						errMsg.append(value.toString()).append("]")
						.append("不能转换成日期,正确的格式应该是:[").append(pattern+"]");
						throw new ExcelDataException(errMsg.toString(), rowNum, fieldValue.getAlias(),value,bean);
					}
					return date;
				} else if (value instanceof Date) {
					return value;
				} else if(value instanceof Number){
					Number val = (Number) value;
					return new Date(val.longValue());
				} else {
					throw new ExcelDataException("数据格式错误,[ " + name + " ]的类型是:" + value.getClass()+",无法转换成日期", rowNum, fieldValue.getAlias(),value,bean);
				}
			}
		} else if (format != null) {
			return resolverExpression(value.toString(), format, type,fieldValue,rowNum,bean);
		} else if (decimalFormat!=null) {
			if (Type.IMPORT == type) {
				if(value instanceof String){
					try {
						return decimalFormat.parse(value.toString());
					} catch (ParseException e) {
						throw new ExcelDataException(e.getMessage(), rowNum, fieldValue.getAlias(),value,bean);
					}
				}
			}else if(Type.EXPORT == type){
				if(value instanceof String){
					value = BeanUtil.convert(value, BigDecimal.class);
				}
				return decimalFormat.format(value);
			}
		} else {
			return value;
		}
		return fieldValue.getDefaultValue();
	}
	
	/**
	 * 日期转换
	 * @param str
	 * @param parsePatterns
	 * @return
	 */
	protected static Date parseDate(String str,String ... parsePatterns){
		if(parsePatterns!= null && parsePatterns.length > 0){
			for (String p : parsePatterns) {
				SimpleDateFormat sdf = new SimpleDateFormat(p);
				try{
					return sdf.parse(str);
				}catch(Exception ignore){
					continue;
				}
			}
		}
		return null;
	}
	
	/**
	 * 解析表达式format 属性
	 * 
	 * @param value
	 * @param format
	 * @param fieldValue
	 * @param rowNum
	 * @return
	 */
	protected String resolverExpression(String value, String format, Type type,FieldValue fieldValue,int rowNum,Object refObject) {
		try {
			String[] expressions = StringUtils.split(format, ",");
			for (String expression : expressions) {
				String[] val = StringUtils.split(expression, ":");
				String v1 = val[0];
				String v2 = val[1];
				if (Type.EXPORT == type) {
					if (value.equals(v1)) {
						return v2;
					}
				} else if (Type.IMPORT == type) {
					if (value.equals(v2)) {
						return v1;
					}
				}
			}
		} catch (Exception e) {
			throw new ExcelDataException("表达式:[" + format + "]错误,正确的格式应该以[,]号分割,[:]号取值", rowNum, fieldValue.getAlias(),value,refObject);
		}
		throw new ExcelDataException("["+value+"]取值错误", rowNum, fieldValue.getAlias(),value,refObject);
	}
	
}
