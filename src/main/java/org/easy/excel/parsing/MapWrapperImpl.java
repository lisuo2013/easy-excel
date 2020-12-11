package org.easy.excel.parsing;

import java.util.Map;

import org.springframework.beans.AbstractPropertyAccessor;
import org.springframework.beans.BeansException;
import org.springframework.core.convert.TypeDescriptor;

/**
 * Map 类型的bean处理
 * @author lisuo
 *
 */
@SuppressWarnings("all")
class MapWrapperImpl extends AbstractPropertyAccessor{
	
	private Map map;
	
	public MapWrapperImpl(Map map) {
		this.map = map;
	}

	@Override
	public boolean isReadableProperty(String propertyName) {
		return true;
	}

	@Override
	public boolean isWritableProperty(String propertyName) {
		return true;
	}

	@Override
	public TypeDescriptor getPropertyTypeDescriptor(String propertyName) throws BeansException {
		return null;
	}

	@Override
	public Object getPropertyValue(String propertyName) throws BeansException {
		return map.get(propertyName);
	}

	@Override
	public void setPropertyValue(String propertyName, Object value) throws BeansException {
		map.put(propertyName, value);
	}

	public Map getRootInstance() {
		return map;
	}
	
	
}
