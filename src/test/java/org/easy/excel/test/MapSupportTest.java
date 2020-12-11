package org.easy.excel.test;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.easy.excel.ExcelContext;
import org.easy.excel.result.ExcelImportResult;
import org.junit.Test;
import org.springframework.core.io.ClassPathResource;

/**
 * Map映射例子
 * 
 * @author lisuo
 *
 */
public class MapSupportTest {
	// 测试时文件磁盘路径
	private static String path = "test-excel.xlsx";
	// 配置文件路径
	private static ExcelContext context = new ExcelContext("excel-config.xml");
	// Excel配置文件中配置的id
	private static String excelId = "studentMap";

	@Test
	public void testImport() throws Exception {
		ClassPathResource resource = new ClassPathResource(path);
		// 第二个参数需要注意,它是指标题索引的位置,可能你的前几行并不是标题,而是其他信息,
		// 比如数据批次号之类的,关于如何转换成javaBean,具体参考配置信息描述
		ExcelImportResult result = context.readExcel(excelId, 2, resource.getInputStream());
		System.out.println(result.getHeader());
		List<Map<String, Object>> stus = result.getListBean();
		for (Map<String, Object> stu : stus) {
			System.out.println(stu);
		}
		resource.getInputStream().close();
		// 这种方式和上面的没有任何区别,底层方法默认标题索引为0
		// context.readExcel(excelId, fis);
	}

	@Test
	public void testExportSimple() throws Exception {
		OutputStream ops = new FileOutputStream("src/test/resources/test-export-excel-map.xlsx");
		Workbook workbook = context.createExcel(excelId, getStudents());
		workbook.write(ops);
		ops.close();
		workbook.close();
	}

	// 获取模拟数据,数据库数据...
	public static List<Map<String, Object>> getStudents() {
		int size = 5;
		List<Map<String, Object>> students = new ArrayList<>(size);
		for (int i = 0; i < size; i++) {
			Map<String, Object> stu = new HashMap<>();
			stu.put("id", "" + (i + 1));
			stu.put("name", "张三" + i);
			stu.put("age", 20);
			stu.put("studentNo", "Stu_" + i);
			stu.put("createTime", new Date());
			stu.put("status", i % 2 == 0 ? 1 : 0);
			stu.put("createUser", "王五" + i);

			// 创建复杂对象
			if (i % 2 == 0) {
				stu.put("book.bookName", "Thinking in java");
				stu.put("book.price", 12345.1253);
				stu.put("book.author.authorName", "Bruce Eckel");
			}
			students.add(stu);
		}
		return students;
	}

}
