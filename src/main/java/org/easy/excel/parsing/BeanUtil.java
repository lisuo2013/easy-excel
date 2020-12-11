package org.easy.excel.parsing;


import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.CollectionUtils;
import org.springframework.beans.AbstractNestablePropertyAccessor;
import org.springframework.beans.AbstractPropertyAccessor;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeanWrapperImpl;
import org.springframework.beans.DirectFieldAccessor;
import org.springframework.beans.NotReadablePropertyException;
import org.springframework.beans.NotWritablePropertyException;
import org.springframework.core.convert.support.DefaultConversionService;
import org.springframework.util.ConcurrentReferenceHashMap;
import org.springframework.util.ConcurrentReferenceHashMap.ReferenceType;

/**
 * 
 * Bean 处理
 * @author lisuo
 */
abstract class BeanUtil {
	
	//spring 类型转换配置
	private static DefaultConversionService defaultConversionService = new DefaultConversionService();
	
	//软引用Map
	private static final Map<Class<?>, List<Field>> declaredFieldsCache = new ConcurrentReferenceHashMap<>(256,ReferenceType.SOFT);
	private static final Map<Class<?>[], Class<?>> eqSuperClassCache = new ConcurrentReferenceHashMap<>(16,ReferenceType.SOFT);
	
	private static final List<Field> NO_FIELDS = Collections.emptyList();
	
	/**
	 * 获取指定类的所有字段,排除static,final字段
	 * @param clazz 类型
	 * @return List<字段>
	 */
	public static List<Field> getFields(Class<?> clazz){
		List<Field> result = declaredFieldsCache.get(clazz);
		if(result==null) {
			Class<?> oldClazz = clazz;
			result = new ArrayList<Field>();
			while(clazz!=Object.class){
				try {
					Field[] fields = clazz.getDeclaredFields();
					for (Field field:fields) {
						int modifiers = field.getModifiers();
						//过滤static或final字段
						if(Modifier.isStatic(modifiers)||Modifier.isFinal(modifiers)){
							continue;
						}
						result.add(field);
					}
				} catch (Exception ignore) {}
				clazz = clazz.getSuperclass();
			}
			declaredFieldsCache.put(oldClazz, (result.isEmpty() ? NO_FIELDS : result));
		}
		return result;
	}
	
	/**
	 * 获取指定类的所有字段名称,排除static,final字段
	 * @param clazz 类型
	 * @return List<字段名称>
	 */
	public static List<String> getFieldNames(Class<?> clazz){
		List<Field> fields = getFields(clazz);
		List<String> fieldNames = new ArrayList<String>(fields.size());
		for(Field field:fields){
			fieldNames.add(field.getName());
		}
		return fieldNames;
	}
	
	/**
	 * 反射无参创建对象
	 * @param clazz
	 * @return
	 */
	public static <T> T newInstance(Class<T> clazz){
		return BeanUtils.instantiateClass(clazz);
	}
	
	/**
	 * 构建属性访问器
	 * @param bean pojo实例
	 * @param nested 是否支持内嵌属性如stu.name或stu.books[0].name
	 * @return 属性访问器
	 */
	@SuppressWarnings("all")
	public static AbstractPropertyAccessor buildAccessor(Object bean, boolean nested){
		if(bean instanceof Map) {
			return new MapWrapperImpl((Map)bean);
		}
		AbstractPropertyAccessor accessor = new BeanWrapperImpl(bean);
		accessor.setAutoGrowNestedPaths(nested);
		accessor.setConversionService(defaultConversionService);
		return accessor;
	}
	
	//构建私有属性的访问器
	private static AbstractPropertyAccessor buildDirectFieldAccessor(Object bean, boolean nested){
		AbstractPropertyAccessor accessor = new DirectFieldAccessor(bean);
		accessor.setAutoGrowNestedPaths(nested);
		accessor.setConversionService(defaultConversionService);
		return accessor;
	}
	
	/**
	 * 设置value 支持私有属性
	 * @param accessor
	 * @param name 字段名称
	 * @param value 值
	 * @param ignoreError 是否忽略找不到属性错误 
	 */
	public static void setPropertyValue(AbstractPropertyAccessor accessor,String name,Object value,boolean ignoreError){
		if(value!=null){
			try{
				accessor.setPropertyValue(name, value);
			}catch(NotWritablePropertyException e){
				if(accessor instanceof BeanWrapperImpl){
					accessor = buildDirectFieldAccessor(((AbstractNestablePropertyAccessor) accessor).getRootInstance(),true);
					try{
						accessor.setPropertyValue(name, value);
					}catch(NotWritablePropertyException ex){
						if(!ignoreError){
							throw ex;
						}
					}
				}
			}
		}
	}
	
	/**
	 * 获取属性value ，支持私有属性
	 * @param accessor 属性访问器
	 * @param name 属性名称
	 * @param ignoreError 是否忽略不存在的属性
	 * @return
	 */
	public static Object getPropertyValue(AbstractPropertyAccessor accessor,String name,boolean ignoreError){
		try{
			Object value = accessor.getPropertyValue(name);
			return value;
		}catch(NotReadablePropertyException e){
			if(accessor instanceof BeanWrapperImpl){
				accessor = buildDirectFieldAccessor(((AbstractNestablePropertyAccessor) accessor).getRootInstance(),true);
				try{
					return accessor.getPropertyValue(name);
				}catch(NotReadablePropertyException ex){
					if(!ignoreError){
						throw ex;
					}
				}
			}
		}
		return null;
	}
	
	/**
	 * 获取n个类,相同的父类类型,如果多个相同的父类,获取最接近的的,
	 * 如果传递的对象包含Object.class 直接返回null 
	 * @param clazzs 
	 * @return 相同的父类Class
	 */
	public static Class<?> getEqSuperClass(Class<?> ...clazzs){
		Class<?> ret = eqSuperClassCache.get(clazzs);
		if(ret == null) {
			ret = Object.class;
			List<List<Class<?>>> container = new ArrayList<List<Class<?>>>(clazzs.length);
			for(Class<?>clazz :clazzs){
				if(clazz==Object.class) {
					return null;
				}
				List<Class<?>> superClazz = new ArrayList<Class<?>>();
				for(clazz=clazz.getSuperclass();clazz!=Object.class;clazz=clazz.getSuperclass()){
					superClazz.add(clazz);
				}
				container.add(superClazz);
			}
			List<Class<?>> result = new ArrayList<Class<?>>();  
			Iterator<List<Class<?>>> it = container.iterator();
			int len = 0;
			while(it.hasNext()){
				if(len == 0){
					result.addAll(it.next());
				}else{
					result.retainAll(it.next());
					if(CollectionUtils.isEmpty(result)){
						break;
					}
				}
				len++;
			}
			//不管相同父类有几个,返回最接近的
			if(!CollectionUtils.isEmpty(result)){
				ret = result.get(0);
			}
			eqSuperClassCache.put(clazzs, ret);
		}
		return ret;
	}
	
	public static <T> T convert(Object source, Class<T> targetType) {
		return defaultConversionService.convert(source, targetType);
	}
}
