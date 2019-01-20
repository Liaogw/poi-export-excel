package com.demo.poi.download.convertFormatImp;

import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

import com.demo.poi.download.XlsConvertFormatFilter;

public class DateFormatConvertionImp implements XlsConvertFormatFilter {

	private static DateTimeFormatter formatter;

	private DateFormatConvertionImp() {
	}

	@Override
	public String convert(String data) {
		if ((data == null) || data.equals("")) {
			return "";
		}
		try {
			LocalDateTime dateTime = new Date(Long.valueOf(data)).toInstant().atZone(ZoneId.systemDefault())
					.toLocalDateTime();
			return dateTime.format(DateFormatConvertionImp.formatter);
		} catch (NumberFormatException e) {
			return data;
		}
	}

	private static XlsConvertFormatFilter filter;

	public static XlsConvertFormatFilter get() {
		formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm");
		if (filter == null) {
			filter = new DateFormatConvertionImp();
		}
		return filter;
	}

	public static XlsConvertFormatFilter get(String format) {
		formatter = DateTimeFormatter.ofPattern(format);
		if (filter == null) {
			filter = new DateFormatConvertionImp();
		}
		return filter;
	}
}
