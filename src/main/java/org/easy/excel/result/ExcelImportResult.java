package org.easy.excel.result;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;
import org.easy.excel.exception.ExcelDataException;

/**
 * Excel导入结果
 * 
 * @author lisuo
 *
 */
public class ExcelImportResult {
	
	/** 头信息,标题行之前的数据,每行表示一个List<Object>,每个Object存放一个cell单元的值 */
	private List<List<Object>> header = null;

	/** JavaBean集合,从标题行下面解析的数据 ,校验通过的数据 */
	private List<?> listBean;
	
	/** Errors */
	private List<ExcelDataException> errors = new ArrayList<ExcelDataException>();
	
	/** Excel中需要处理的数据量,假设10条数据,8条校验未通过,这个值是10,listBean是2条数据 */
	private Integer totalNum;
	
	public List<List<Object>> getHeader() {
		return header;
	}

	public void setHeader(List<List<Object>> header) {
		this.header = header;
	}
	
	@SuppressWarnings("unchecked")
	public <T> List<T> getListBean() {
		return (List<T>) listBean;
	}

	public void setListBean(List<?> listBean) {
		this.listBean = listBean;
	}
	
	public List<ExcelDataException> getErrors() {
		return errors;
	}
	
	public Integer getTotalNum() {
		return totalNum;
	}
	
	public void setTotalNum(Integer totalNum) {
		this.totalNum = totalNum;
	}
	
	/**
	 * 导入是否含有错误
	 * @return true:有错误,false:没有错误
	 */
	public boolean hasErrors(){
		return CollectionUtils.isNotEmpty(errors);
	}
	
}
