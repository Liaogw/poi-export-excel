package com.demo.poi.download;

import com.alibaba.fastjson.JSONObject;

public interface XlsConvertDataFilter {

	/**
	 * 从 obj 中取出 order 类型的对象
	 * 
	 * @param obj
	 * @param order
	 * @return
	 */
	public String convert(JSONObject obj, String order);
}
