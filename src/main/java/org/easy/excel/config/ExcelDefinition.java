package org.easy.excel.config;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * Excel定义
 * 
 * @author lisuo
 *
 */
public class ExcelDefinition {

	/** ID,必须 */
	private String id;

	/** 全类名,必须 */
	private String className;

	/** Class信息 */
	private Class<?> clazz;

	/**导出时,sheet页的名称,可以不设置*/
	private String sheetname;
	
	/**导出时,sheet页所有的默认列宽,可以不设置*/
	private Integer defaultColumnWidth;

	/**导出时,cell默认对其方式:支持,center,left,right*/
	private HorizontalAlignment defaultAlign;
	
	/** Field属性的全部定义 */
	private List<FieldValue> fieldValues = new ArrayList<FieldValue>();
	
	/** Excel 文件sheet索引，默认为0即，第一个 */
	private int sheetIndex = 0;
	
	/** 不能为空的数据是否在标题前生成*（红色），以及在导入时标题前缀加*处理 */
	private boolean requiredTag;

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getClassName() {
		return className;
	}

	public void setClassName(String className) {
		this.className = className;
	}

	public Class<?> getClazz() {
		return clazz;
	}

	public void setClazz(Class<?> clazz) {
		this.clazz = clazz;
	}

	public List<FieldValue> getFieldValues() {
		return fieldValues;
	}

	public void setFieldValues(List<FieldValue> fieldValues) {
		this.fieldValues = fieldValues;
	}

	public String getSheetname() {
		return sheetname;
	}
	
	public void setSheetname(String sheetname) {
		this.sheetname = sheetname;
	}

	public Integer getDefaultColumnWidth() {
		return defaultColumnWidth;
	}

	public void setDefaultColumnWidth(Integer defaultColumnWidth) {
		this.defaultColumnWidth = defaultColumnWidth;
	}

	public HorizontalAlignment getDefaultAlign() {
		return defaultAlign;
	}

	public void setDefaultAlign(HorizontalAlignment defaultAlign) {
		this.defaultAlign = defaultAlign;
	}
	
	public int getSheetIndex() {
		return sheetIndex;
	}
	
	public void setSheetIndex(int sheetIndex) {
		this.sheetIndex = sheetIndex;
	}
	
	public boolean isRequiredTag() {
		return requiredTag;
	}
	
	public void setRequiredTag(boolean requiredTag) {
		this.requiredTag = requiredTag;
	}
	
}
