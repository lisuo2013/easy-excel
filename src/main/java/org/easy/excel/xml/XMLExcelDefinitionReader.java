package org.easy.excel.xml;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.StringWriter;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.easy.excel.ExcelDefinitionReader;
import org.easy.excel.config.ExcelDefinition;
import org.easy.excel.config.FieldValue;
import org.easy.excel.exception.ExcelException;
import org.easy.excel.parsing.CellValueConverter;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.util.Assert;
import org.springframework.util.ResourceUtils;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * Excel XML定义读取注册
 * @author lisuo
 *
 */
public class XMLExcelDefinitionReader implements ExcelDefinitionReader{
	
	/** 注册信息 */
	private Map<String, ExcelDefinition> registry;
	
	/** 配置文件路径 */
	private String locations;
	
	/**
	 * @param location xml配置路径
	 */
	@SuppressWarnings("resource")
	public XMLExcelDefinitionReader(String locations) {
		try{
			if(StringUtils.isBlank(locations)){
				throw new ExcelException("locations 不能为空");
			}
			this.locations = locations;
			registry = new HashMap<String, ExcelDefinition>();
			String[] locationArr = StringUtils.split(locations, ",");
			for (String location:locationArr) {
				InputStream fis = null;
				try{
					File file = ResourceUtils.getFile(location);
					fis = new FileInputStream(file);
				}catch(FileNotFoundException e){
					//如果没有找到文件,默认尝试从类路径加载
					Resource resource = new ClassPathResource(location);
					fis = resource.getInputStream();
				}
				loadExcelDefinitions(fis);
			}
		}catch(ExcelException e){
			throw e;
		}catch(Exception e){
			throw new ExcelException(e);
		}
	}
	
	/**
	 * 加载解析配置文件信息
	 * @param inputStream
	 * @throws Exception
	 */
	protected void loadExcelDefinitions(InputStream inputStream) throws Exception {
		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		//不校验DTD约束文档
		factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false);
		DocumentBuilder docBuilder = factory.newDocumentBuilder();
		Document doc = docBuilder.parse(inputStream);
		registerExcelDefinitions(doc);
		inputStream.close();
	}

	/**
	 * 注册Excel定义信息
	 * @param doc
	 */
	protected void registerExcelDefinitions(Document doc) {
		Element root = doc.getDocumentElement();
		NodeList nl = root.getChildNodes();
		for (int i = 0; i < nl.getLength(); i++) {
			Node node = nl.item(i);
			if (node instanceof Element) {
				Element ele = (Element) node;
				processExcelDefinition(ele);
			}
		}
	}

	/**
	 * 解析和校验Excel定义
	 * @param ele
	 */
	protected void processExcelDefinition(Element ele) {
		ExcelDefinition excelDefinition = new ExcelDefinition();
		String id = ele.getAttribute("id");
		Validate.notNull(id, "Excel 配置文件[" + locations + "] , id为 [ null ] ");
		if (registry.containsKey(id)) {
			throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [" + id + "] 的不止一个");
		}
		excelDefinition.setId(id);
		String className = ele.getAttribute("class");
		try{
			Validate.notNull(className, "Excel 配置文件[" + locations + "] , id为 [" + id + "] 的 class 为 [ null ]");
		}catch(Exception e){
			throw new ExcelException(e.getMessage());
		}
		Class<?> clazz = null;
		try {
			if(className.toLowerCase().equals("map") || className.toLowerCase().equals("hashmap")) {
				className = "java.util.HashMap";
			}
			clazz = Class.forName(className);
		} catch (ClassNotFoundException e) {
			throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [" + id + "] 的 class 为 [" + className + "] 类不存在 ");
		}
		excelDefinition.setClassName(className);
		excelDefinition.setClazz(clazz);
		if(StringUtils.isNotBlank(ele.getAttribute("defaultColumnWidth"))){
			try{
				excelDefinition.setDefaultColumnWidth(Integer.parseInt(ele.getAttribute("defaultColumnWidth")));
			}catch(NumberFormatException e){
				throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
				+ " ] 的 defaultColumnWidth 属性不能为 [ "+ele.getAttribute("defaultColumnWidth")+" ],只能为int类型");
			}
		}
		if(StringUtils.isNotBlank(ele.getAttribute("sheetname"))){
			excelDefinition.setSheetname(ele.getAttribute("sheetname"));
		}
		//默认对齐方式
		String defaultAlign = ele.getAttribute("defaultAlign");
		if(StringUtils.isNotBlank(defaultAlign)){
			try{
				//获取cell对齐方式的常量值
				excelDefinition.setDefaultAlign(HorizontalAlignment.valueOf(defaultAlign.toUpperCase()));
			}catch(Exception e){
				throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
				+ " ] 的 defaultAlign 属性不能为 [ "+defaultAlign+" ],目前支持的["+Arrays.asList(HorizontalAlignment.values())+"]");
			}
		}
		
		//requiredTag处理
		String requiredTag = ele.getAttribute("requiredTag");
		if(StringUtils.isNotBlank(requiredTag)){
			excelDefinition.setRequiredTag(Boolean.parseBoolean(requiredTag));
		}
		
		//Sheet Index
		if(StringUtils.isNotBlank(ele.getAttribute("sheetIndex"))){
			try{
				int sheetIndex = Integer.parseInt(ele.getAttribute("sheetIndex"));
				if(sheetIndex<0){
					throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
					+ " ] 的 sheetIndex 属性不能为 [ "+ele.getAttribute("sheetIndex")+" ],索引位置至少从0开始");
				}
				excelDefinition.setSheetIndex(sheetIndex);
			}catch(NumberFormatException e){
				throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
				+ " ] 的 sheetIndex 属性不能为 [ "+ele.getAttribute("sheetIndex")+" ],只能为int类型");
			}
		}
		processField(ele, excelDefinition);
		registry.put(id, excelDefinition);
	}
	
	/**
	 * 解析和校验Field属性定义
	 * @param ele
	 * @param excelDefinition
	 */
	protected void processField(Element ele, ExcelDefinition excelDefinition) {
		NodeList fieldNode = ele.getElementsByTagName("field");
		if (fieldNode != null) {
			for (int i = 0; i < fieldNode.getLength(); i++) {
				Node node = fieldNode.item(i);
				if (node instanceof Element) {
					FieldValue fieldValue = new FieldValue();
					Element fieldEle = (Element) node;
					String name = fieldEle.getAttribute("name");
					Validate.isTrue(StringUtils.isNotBlank(name), "Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
							+ " ] 的 name 属性不能为 [ null ]");
					fieldValue.setName(name);
					String pattern = fieldEle.getAttribute("pattern");
					if(StringUtils.isNotBlank(pattern)){
						fieldValue.setPattern(pattern);
					}
					String format = fieldEle.getAttribute("format");
					if(StringUtils.isNotBlank(format)){
						fieldValue.setFormat(format);;
					}
					String isNull = fieldEle.getAttribute("isNull");
					if(StringUtils.isNotBlank(isNull)){
						fieldValue.setNull(Boolean.parseBoolean(isNull));
					}
					String regex = fieldEle.getAttribute("regex");
					if(StringUtils.isNotBlank(regex)){
						fieldValue.setRegex(unescapeRegex(regex));
					}
					String regexErrMsg = fieldEle.getAttribute("regexErrMsg");
					if(StringUtils.isNotBlank(regexErrMsg)){
						fieldValue.setRegexErrMsg(regexErrMsg);
					}
					
					//标题设置
					String title = fieldEle.getAttribute("title");
					try{
						Assert.isTrue(StringUtils.isNotBlank(title), "Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
						+ " ] 的 title 属性不能为 [ null ]");
					}catch(Exception e){
						throw new ExcelException(e.getMessage());
					}
					boolean requiredTag = excelDefinition.isRequiredTag();
					if(requiredTag){
						if(!title.startsWith("*")){
							if(!fieldValue.isNull()){
								title = "*"+title;
							}
						}
					}
					//别名去除*号
					if(title.startsWith("*")){
						fieldValue.setAlias(title.substring(1));
					}else{
						fieldValue.setAlias(title);
					}
					fieldValue.setTitle(title);
					//对齐方式
					String align = fieldEle.getAttribute("align");
					if(StringUtils.isNotBlank(align)){
						try{
							fieldValue.setAlign(HorizontalAlignment.valueOf(align.toUpperCase()));
						}catch(Exception e){
							throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
							+ " ] 的 align 属性不能为 [ "+align+" ],目前支持的["+Arrays.asList(HorizontalAlignment.values())+"]");
						}
					}
					//cell 宽度
					String columnWidth = fieldEle.getAttribute("columnWidth");
					if(StringUtils.isNotBlank(columnWidth)){
						try{
							int intVal = Integer.parseInt(columnWidth);
							fieldValue.setColumnWidth(intVal);
						}catch(NumberFormatException e){
							throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
							+ " ] 的 columnWidth 属性 [ "+columnWidth+" ] 不是一个合法的数值");
						}
					}
					//cell标题背景色
					String titleBgColor = fieldEle.getAttribute("titleBgColor");
					if(StringUtils.isNotBlank(titleBgColor)){
						try{
							fieldValue.setTitleBgColor(IndexedColors.valueOf(titleBgColor.toUpperCase()).index);
						}catch(Exception e){
							throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
							+ " ] 的 titleBgColor 属性不能为 [ "+titleBgColor+" ],支持的颜色有["+Arrays.asList(IndexedColors.values())+"]");
						}
					}
					//cell标题字体颜色
					String titleFountColor = fieldEle.getAttribute("titleFountColor");
					if(StringUtils.isNotBlank(titleFountColor)){
						try{
							fieldValue.setTitleFountColor(IndexedColors.valueOf(titleFountColor.toUpperCase()).index);
						}catch(Exception e){
							throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
							+ " ] 的 titleFountColor 属性不能为 [ "+titleFountColor+" ],支持的颜色有["+Arrays.asList(IndexedColors.values())+"]");
						}
					}
					//cell 样式是否与标题样式一致
					String uniformStyle = fieldEle.getAttribute("uniformStyle");
					if(StringUtils.isNotBlank(uniformStyle)){
						fieldValue.setUniformStyle(Boolean.parseBoolean(uniformStyle));
					}
					
					//解析自定义转换器
					String cellValueConverterName = fieldEle.getAttribute("cellValueConverter");
					if(StringUtils.isNotBlank(cellValueConverterName)){
						try {
							Class<?> clazz = Class.forName(cellValueConverterName);
							if(!CellValueConverter.class.isAssignableFrom(clazz)){
								throw new ExcelException("配置的："+cellValueConverterName+"错误,不是一个标准的["+CellValueConverter.class.getName()+"]实现");
							}
							fieldValue.setCellValueConverterName(cellValueConverterName);
						} catch (ClassNotFoundException e) {
							throw new ExcelException("无法找到定义的解析器：["+cellValueConverterName+"]"+"请检查配置信息");
						}
					}
					
					//roundingMode 解析
					String roundingMode = fieldEle.getAttribute("roundingMode");
					if(StringUtils.isNotBlank(roundingMode)){
						try{
							//获取roundingMode常量值
							fieldValue.setRoundingMode(RoundingMode.valueOf(roundingMode.toUpperCase()));
						}catch(Exception e){
							throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
							+ " ] 的 roundingMode 属性不能为 [ "+roundingMode+" ],具体看[java.math.RoundingMode]支持的值");
						}
					}
					
					//解析decimalFormat
					String decimalFormatPattern = fieldEle.getAttribute("decimalFormatPattern");
					if(StringUtils.isNotBlank(decimalFormatPattern)){
						try{
							fieldValue.setDecimalFormatPattern(decimalFormatPattern);
							fieldValue.setDecimalFormat(new DecimalFormat(decimalFormatPattern));
							fieldValue.getDecimalFormat().setRoundingMode(fieldValue.getRoundingMode());
						}catch(Exception e){
							throw new ExcelException("Excel 配置文件[" + locations + "] , id为 [ " + excelDefinition.getId()
							+ " ] 的 decimalFormatPattern 属性不能为 [ "+decimalFormatPattern+" ],请配置标准的JAVA格式");
						}
					}

					//解析其他配置项
					String otherConfig = fieldEle.getAttribute("otherConfig");
					if(StringUtils.isNotBlank(otherConfig)){
						fieldValue.setOtherConfig(otherConfig);
					}
					
					//解析,值为空时,字段的默认值
					String defaultValue = fieldEle.getAttribute("defaultValue");
					if(StringUtils.isNotBlank(defaultValue)){
						fieldValue.setDefaultValue(defaultValue);
					}
					
					//处理forceText
					fieldValue.setForceText(Boolean.parseBoolean(fieldEle.getAttribute("forceText")));
					
					excelDefinition.getFieldValues().add(fieldValue);
				}
			}
		}
	}
	
	/**
	 * 处理正则表达式写法的问题：把下列两种写法转义成统一的去除多余的\\符号，导致正则匹配错误
	 * ^[1-9]\d*$ 标准的正则表达式，没有java的转义
	 * 下面的方式使用了java转义，正则在运行时会导致运行错误
	 * ^[1-9]\\d*$
	 * @param str
	 * @return
	 */
	protected String unescapeRegex(String str) {
		StringWriter out = new StringWriter(str.length());
		int sz = str.length();
		int pre = 0;
		for (int i = 0; i < sz; i++) {
			char ch = str.charAt(i);
			switch (ch) {
			case '\\':
				//如果上次和当前写入是连续的，那么本次不再写入
				if (pre + 1 == i) {
					break;
				} else {
					out.write('\\');
					pre = i;
				}
				break;
			default:
				out.write(ch);
				break;
			}
		}
		return out.toString();
	}

	/**
	 * @return 不可被修改的注册信息
	 */
	@Override
	public Map<String, ExcelDefinition> getRegistry() {
		return Collections.unmodifiableMap(registry);
	}
	
	public String getLocations() {
		return locations;
	}
}
