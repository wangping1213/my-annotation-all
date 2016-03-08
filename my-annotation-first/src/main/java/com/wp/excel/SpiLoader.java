package com.wp.excel;

import java.util.Iterator;
import java.util.ServiceLoader;

/**
 * spi接口调用加载器
 * @author wangping
 * @time 2016年1月16日 下午3:47:50
 */
public class SpiLoader {
	
	/**
	 * 根据对应类型的class来取得对应对象（最多只返回一个对象）
	 * @author wangping
	 * @time 2016年1月16日 下午3:42:42
	 * @param eClass
	 * @return
	 */
	public static <E> E getLoader(Class<E> eClass) {
		E e = null;
		ServiceLoader<E> eLoader = ServiceLoader.load(eClass);
		Iterator<E> eIterator = eLoader.iterator();
		if (eIterator.hasNext()) {
			e = eIterator.next();
		}
		return e;
	}

}
