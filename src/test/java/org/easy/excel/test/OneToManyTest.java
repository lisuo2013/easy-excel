package org.easy.excel.test;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.easy.excel.ExcelContext;
import org.easy.excel.result.ExcelImportResult;
import org.easy.excel.test.model.BookModel;
import org.easy.excel.test.model.OneToManyModel;
import org.junit.Test;
import org.springframework.core.io.ClassPathResource;

/**
 * 一对多测试用例
 * @author lisuo
 *
 */
public class OneToManyTest {
	
	// 配置文件路径
	private static ExcelContext context = new ExcelContext("excel-config.xml");
	// Excel配置文件中配置的id
	private static String excelId = "oneToManyModel";
	
	
	/**
	 * 导入Excel,使用了org.easy.excel.test.ExportTest.testExportCustomHeader()方法生成的Excel
	 * @throws Exception
	 */
	@Test
	public void testImport()throws Exception{
		ClassPathResource resource = new ClassPathResource("OneToManyTest-excel.xlsx");
		//第二个参数需要注意,它是指标题索引的位置,可能你的前几行并不是标题,而是其他信息,
		//比如数据批次号之类的,关于如何转换成javaBean,具体参考配置信息描述
		ExcelImportResult result = context.readExcel(excelId, 0, resource.getInputStream());
		System.out.println(result.getHeader());
		List<OneToManyModel> stus = result.getListBean();
		for(OneToManyModel one:stus){
			System.out.println(one);
		}
		resource.getInputStream().close();
	}
	
	@Test
	public void testExport()throws Exception{
		OutputStream ops = new FileOutputStream("src/test/resources/OneToManyTest-excel.xlsx");
		Workbook workbook = context.createExcel(excelId,getOneToManys());
		workbook.write(ops);
		ops.close();
		workbook.close();
	}
	
	public List<OneToManyModel> getOneToManys(){
		List<OneToManyModel> list = new ArrayList<OneToManyModel>();
		for (int i = 1; i <= 10; i++) {
			OneToManyModel one = new OneToManyModel();
			one.setStudentName("张三"+i);
			List<BookModel> books = new ArrayList<BookModel>();
			for (int j = 0; j <2; j++) {
				BookModel book = new BookModel();
				book.setBookName("Hello"+j);
				books.add(book);
			}
			one.setBooks(books);
			list.add(one);
		}
		
		return list;
	}
	
}
