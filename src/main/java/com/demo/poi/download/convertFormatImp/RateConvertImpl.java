package com.demo.poi.download.convertFormatImp;

import com.demo.poi.download.XlsConvertFormatFilter;

/**
 * 同比，环比，转换显示
 */
public class RateConvertImpl implements XlsConvertFormatFilter {

	@Override
	public String convert(String data) {

		if (data == null) {
			return "--";
		} else {
			Double d = Double.valueOf(data) * 100;
			String s = String.format("%.2f", d);
			return s + "%";
		}
	}

	private static XlsConvertFormatFilter filter;

	public static XlsConvertFormatFilter get() {
		if (filter == null) {
			filter = new RateConvertImpl();
		}
		return filter;
	}
}
