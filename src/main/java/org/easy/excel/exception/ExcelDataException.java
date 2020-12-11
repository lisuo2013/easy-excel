package org.easy.excel.exception;

/**
 * excel 数据异常（数据导入校验不通过抛出，在解析数据时）
 * @author lisuo
 *
 */
public class ExcelDataException extends ExcelException {

	/**
	 * 
	 */
	private static final long serialVersionUID = 3527640098624854018L;

	/**
	 * 行号
	 */
	private int row;

	/**
	 * 标题名称
	 */
	private String title;
	/**
	 * 原始value
	 */
	private Object originalValue;
	/**
	 * 错误信息
	 */
	private String errInfo;
	/**
	 * 持有的Bean实例
	 */
	private Object refObject;
	
	

	public ExcelDataException(String message, int row, String title,Object originalValue,Object refObject) {
		super(wholeMessage(message, row, title));
		this.errInfo = message;
		this.row = row;
		this.title = title;
		this.originalValue = originalValue;
		this.refObject = refObject;
	}

	public int getRow() {
		return row;
	}

	public String getTitle() {
		return title;
	}

	public String getErrInfo() {
		return errInfo;
	}
	
	public Object getOriginalValue() {
		return originalValue;
	}
	
	public Object getRefObject() {
		return refObject;
	}

	/**
	 * 获取完整的错误提示信息[行,标题,错误信息]
	 * @return 完整的错误提示信息
	 */
	private static String wholeMessage(String message, int row, String title) {
		return new StringBuilder()
				.append("第[")
				.append(row)
				.append("行],[")
				.append(title)
				.append("]")
				.append(message)
				.toString();
	}

	@Override
	public String toString() {
		return "ExcelDataException [row=" + row + ", title=" + title + ", originalValue=" + originalValue + ", errInfo="
				+ errInfo + ", refObject=" + refObject + "]";
	}

	
	
}
