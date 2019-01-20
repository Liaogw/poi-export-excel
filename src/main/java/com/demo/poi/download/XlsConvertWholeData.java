package com.demo.poi.download;

import java.util.List;

public interface XlsConvertWholeData {
	public String[][] convertData(List<?> data, String[] orders);
}
