package org.easy.excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.easy.excel.exception.ExcelException;
import org.springframework.core.io.ClassPathResource;

/**
 * Excel 操作工具类
 * @author lisuo
 *
 */
public abstract class ExcelUtil {
	
	/**
	 * 读取Excel,支持任何不规则的Excel文件,
	 * 外层List表示所有的数据行，内层List表示每行中的cell单元数据位置
	 * 假设获取一个Excel第三行第二个单元格的数据，例子代码：
	 * FileInputStream excelStream = new FileInputStream(path);
	 * List<List<Object>> list = ExcelUtil.readExcel(excelStream);
	 * System.out.println(list.get(2).get(1));//第三行第二列,索引行位置是2,列的索引位置是1
	 * @param excelStream Excel文件流
	 * @param sheetIndex Excel-Sheet 的索引
	 * @return List<List<Object>> 
	 * @throws Exception
	 */
	public static List<List<Object>> readExcel(InputStream excelStream,int sheetIndex){
		List<List<Object>> datas = new ArrayList<List<Object>>();
		Workbook workbook = getWorkBookByStream(excelStream);
		//只读取第一个sheet
		Sheet sheet = getSheetAt(workbook, sheetIndex);
		int rows = sheet.getPhysicalNumberOfRows();
		for (int i = 0; i < rows; i++) {
			Row row = sheet.getRow(i);
			if(row==null){
				continue;
			}
			short cellNum = row.getLastCellNum();
			List<Object> item = new ArrayList<Object>(cellNum);
			for(int j=0;j<cellNum;j++){
				Cell cell = row.getCell(j);
				Object value = ExcelUtil.getCellValue(cell);
				item.add(value);
			}
			datas.add(item);
		}
		return datas;
	}
	
	/**
	 * 读取Excel,支持任何不规则的Excel文件,默认读取第一个sheet页
	 * 外层List表示所有的数据行，内层List表示每行中的cell单元数据位置
	 * 假设获取一个Excel第三行第二个单元格的数据，例子代码：
	 * FileInputStream excelStream = new FileInputStream(path);
	 * List<List<Object>> list = ExcelUtil.readExcel(excelStream);
	 * System.out.println(list.get(2).get(1));//第三行第二列,索引行位置是2,列的索引位置是1
	 * @param excelStream Excel文件流
	 * @return List<List<Object>> 
	 * @throws Exception
	 */
	public static List<List<Object>> readExcel(InputStream excelStream)throws Exception {
		return readExcel(excelStream,0);
	}
	
	/**
	 * 设置Cell单元的值
	 * 
	 * @param cell
	 * @param value
	 */
	public static void setCellValue(Cell cell, Object value) {
		if (value != null) {
			if (value instanceof String) {
				cell.setCellValue((String) value);
			} else if (value instanceof Number) {
				cell.setCellValue(Double.parseDouble(String.valueOf(value)));
			} else if (value instanceof Boolean) {
				cell.setCellValue((Boolean) value);
			} else if (value instanceof Date) {
				cell.setCellValue((Date) value);
			}else if(value instanceof RichTextString){
				cell.setCellValue((RichTextString)value);
			} else {
				cell.setCellValue(value.toString());
			}
		}
	}
	
	/**
	 * 获取cell值
	 * 
	 * @param cell
	 * @return
	 */
	public static Object getCellValue(Cell cell) {
		Object value = null;
		if (null != cell) {
			switch (cell.getCellType()) {
			// 空白
			case BLANK:
				break;
			// Boolean
			case BOOLEAN:
				value = cell.getBooleanCellValue();
				break;
			// 错误格式
			case ERROR:
				break;
			// 公式
			case FORMULA:
				Workbook wb = cell.getSheet().getWorkbook();
				CreationHelper crateHelper = wb.getCreationHelper();
				FormulaEvaluator evaluator = crateHelper.createFormulaEvaluator();
				value = getCellValue(evaluator.evaluateInCell(cell));
				break;
			// 数值
			case NUMERIC:
				// 处理日期格式
				if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
					value = cell.getDateCellValue();
				} else {
					value = cell.getNumericCellValue();
				}
				break;
			case STRING:
				value = cell.getStringCellValue();
				break;
			default:
				value = null;
			}
		}
		return value;
	}
	
	public static Sheet getSheetAt(Workbook workbook,int sheetIndex){
		try{
			return workbook.getSheetAt(sheetIndex);
		}catch(Exception e){
			throw new ExcelException("找不到对应的sheet页");
		}
	}
	
	/**
	 * 通过文件流获取workbook实例
	 * @param excelStream excel文件流
	 * @return Workbook
	 * @throws Exception
	 */
	public static Workbook getWorkBookByStream(InputStream excelStream){
		try{
			return WorkbookFactory.create(excelStream);
		}catch(EncryptedDocumentException e){
			if(excelStream!=null){
				try{
					//文件格式异常POI框架不会关闭资源导致文件一直被当前应用占用，释放文件资源
					excelStream.close();
				}catch(Exception ignore){}
			}
			throw new ExcelException("导入的文件不是Excel,无法操作");
		}catch(Exception e){
			throw new ExcelException(e);
		}
	}
	
	/**
	 * 通过路径获取workbook实例
	 * @param path 路径：支持，http、classpath，绝对路径
	 * @return
	 * @throws Exception
	 */
	public static Workbook getWorkBookByPath(String path){
		if(StringUtils.isBlank(path)){
			return null;
		}
		try{
			if (path.startsWith("http:")) {
				URL url = new URL(path);
				return getWorkBookByStream(url.openStream());
			}else if(path.startsWith("classpath:")){
				ClassPathResource resource = new ClassPathResource(path.replace("classpath:", ""));
				return getWorkBookByStream(resource.getInputStream());
			}else{
				return getWorkBookByStream(new FileInputStream(new File(path)));
			}
		}catch(Exception e){
			if(e instanceof ExcelException){
				throw (ExcelException)e;
			}
			throw new ExcelException(e);
		}
	}
	
	/**
	 * 向指定的位置写入数据
	 * @param workbook excel
	 * @param sheetIndex sheet索引位
	 * @param rownum 行
	 * @param cellnum 列
	 * @param value 值
	 * @param append 是否追加
	 * @param symbol 追加数据使用的符号
	 */
	public static void write(Workbook workbook,int sheetIndex,int rownum,int cellnum,String value,boolean append,String symbol){
		Cell cell = getSheetAt(workbook, sheetIndex).getRow(rownum).getCell(cellnum);
		if(cell==null){
			cell = getSheetAt(workbook, sheetIndex).getRow(rownum).createCell(cellnum);
		}
		String val = value;
		if(append){
			Object cellValue = getCellValue(cell);
			if(cellValue!=null){
				if(symbol!=null){
					val = cellValue.toString()+symbol+value;
				}else{
					val = cellValue.toString()+value;
				}
			}
		}
		setCellValue(cell, val);
	}
	
}
