package com.demo.poi.download;

import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

public class XlsDownloadFilter {

	public static final int PAGE_SIZE = 1;
	public static final int ROW_COUNT = 65535;
	private static final int MAX_COLUMN_WIDTH = 40;
	private static final int MIN_COLUMN_WIDTH = 8;

	private List<?> data;

	private String field;

	private String fileName;

	private HttpServletResponse response;

	private Map<String, XlsConvertFormatFilter> convertFormatMap = new HashMap<String, XlsConvertFormatFilter>();

	private Map<String, XlsConvertDataFilter> convertDataMap = new HashMap<String, XlsConvertDataFilter>();

	private XlsConvertWholeData converter;

	public XlsDownloadFilter setFileName(String fileName) {
		this.fileName = fileName;
		return this;
	}

	public XlsDownloadFilter setConvertFormat(String field, XlsConvertFormatFilter filter) {
		convertFormatMap.put(field, filter);
		return this;
	}

	public XlsDownloadFilter setConvertData(String field, XlsConvertDataFilter filter) {
		convertDataMap.put(field, filter);
		return this;
	}

	public XlsDownloadFilter setConverter(XlsConvertWholeData converter) {
		this.converter = converter;
		return this;
	}

	public XlsDownloadFilter(List<?> data, String field, HttpServletResponse response) {
		this(data, field, null, response);
	}

	public XlsDownloadFilter(List<?> data, String field, Map<String, XlsConvertFormatFilter> convertFormatMap,
			HttpServletResponse response) {
		this(data, field, convertFormatMap, null, response);
	}

	public XlsDownloadFilter(List<?> data, String field, Map<String, XlsConvertFormatFilter> convertFormatMap,
			Map<String, XlsConvertDataFilter> convertDataMap, HttpServletResponse response) {
		this(data, field, null, convertFormatMap, convertDataMap, response);
	}

	public XlsDownloadFilter(List<?> data, String field, String fileName,
			Map<String, XlsConvertFormatFilter> convertFormatMap, Map<String, XlsConvertDataFilter> convertDataMap,
			HttpServletResponse response) {

		this.converter = null;
		this.fileName = fileName;
		this.data = data;
		this.field = field;
		this.convertFormatMap = convertFormatMap;
		this.convertDataMap = convertDataMap;
		this.response = response;
		if (convertFormatMap == null) {
			this.convertFormatMap = new HashMap<String, XlsConvertFormatFilter>();
		}
		if (convertDataMap == null) {
			this.convertDataMap = new HashMap<String, XlsConvertDataFilter>();
		}
	}

	public void start() {
		JSONObject obj = JSONObject.parseObject(field);
		String[] headers = jsonArray2StringArray(obj.getJSONArray("headers"));
		String[] orders = jsonArray2StringArray(obj.getJSONArray("headersOrder"));
		String[][] arrData = convertData(data, orders);
		arrData = changeFormat(arrData, orders);
		save(this.fileName, headers, arrData, response);
	}

	// 必要时可以被重载
	private String[][] convertData(List<?> data, String[] orders) {
		// 调用传入的convert
		if (this.converter != null) {
			return converter.convertData(data, orders);
		}
		// 调用系统默认的转换
		String[][] newData = new String[data.size()][orders.length];
		Set<String> keys = convertDataMap.keySet();
		for (int i = 0, len = data.size(); i < len; i++) {
			JSONObject obj = (JSONObject) JSONObject.toJSON(data.get(i));
			for (int j = 0; j < orders.length; j++) {
				try {
					if (contains(keys, orders[j])) { // 需要特殊处理的点
						newData[i][j] = convertDataMap.get(orders[j]).convert(obj, orders[j]);
					} else { // 正常解析的点
						if (orders[j].split("\\.").length == 1) {
							newData[i][j] = obj.getString(orders[j]);
						} else {
							String[] order = orders[j].split("\\.");
							JSONObject object = obj;
							for (int k = 0; k < (order.length - 1); k++) {
								if (object == null) {
									newData[i][j] = "";
									continue;
								}
								object = object.getJSONObject(order[k]);
							}
							newData[i][j] = object.getString(order[order.length - 1]);
						}
					}

				} catch (Throwable e) {
					e.printStackTrace();
					newData[i][j] = "";
				}
			}
		}
		return newData;
	}

	private String[][] changeFormat(String[][] arrData, String[] orders) {
		Object[] keys = convertFormatMap.keySet().toArray();
		int[] positions = new int[keys.length];

		for (int i = 0; i < keys.length; i++) {
			for (int j = 0; j < orders.length; j++) {
				if (keys[i].toString().equalsIgnoreCase(orders[j])) {
					positions[i] = j;
					break;
				}
			}
		}
		for (Integer i = 0; i < positions.length; i++) {
			for (String[] data : arrData) {
				try {
					data[positions[i]] = convertFormatMap.get(keys[i]).convert(data[positions[i]]);
				} catch (Exception e) {
					data[positions[i]] = "";
				}
			}
		}
		return arrData;
	}

	private void save(String name, String[] headers, String[][] arrData, HttpServletResponse response) {
		if (name == null) {
			name = "文件-";
		}
		Workbook workbook = createXls(headers, arrData, response);
		// 添加将数据进行转换
		SimpleDateFormat dateFormater = new SimpleDateFormat("yyyyMMddHHmmss");
		Date date = new Date();
		response.reset();
		response.setContentType("application/vnd.ms-excel");

		String fileName = name + "-" + dateFormater.format(date) + ".xls";
		try {
			response.setHeader("Content-disposition",
					"attachment;filename=" + new String(fileName.getBytes("gbk"), "iso-8859-1"));
		} catch (UnsupportedEncodingException e1) {
			e1.printStackTrace();
		}

		try (OutputStream out = response.getOutputStream();) {
			workbook.write(out);
			out.flush();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static Workbook createXls(String[] headers, String[][] content, HttpServletResponse response) {

		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		CellStyle headCellStyle = workbook.createCellStyle();
		headCellStyle.setAlignment(HorizontalAlignment.CENTER);
		Font font = workbook.createFont();
		headCellStyle.setFont(font);

		int[] width = getColumnWidth(headers, content);
		for (int i = 0; i < headers.length; i++) {
			sheet.setColumnWidth(i, width[i] * 256);
		}

		Row row = sheet.createRow(0);
		writeToRow(headers, row, headCellStyle);

		// 写入需要的数据
		CellStyle bodyCellStyle = workbook.createCellStyle();
		for (int i = 0; i < content.length; i++) {
			row = sheet.createRow(i + 1);
			writeToRow(content[i], row, bodyCellStyle);
		}
		return workbook;
	}

	private static void writeToRow(String[] data, Row row, CellStyle style) {
		for (int i = 0; i < data.length; i++) {
			Object value = data[i];
			if (data[i] == null) {
				value = "";
			}
			Cell cell = row.createCell(i);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			cell.setCellStyle(style);
			cell.setCellValue(String.valueOf(value));
		}
	}

	private String[] jsonArray2StringArray(JSONArray arr) {
		String[] str = new String[arr.size()];
		for (int i = 0; i < arr.size(); i++) {
			str[i] = arr.getString(i);
		}
		return str;
	}

	private static int[] getColumnWidth(String[] headers, String[][] data) {
		int[] width = new int[headers.length];
		int detectLen = 10 > data.length ? data.length : 10;
		int len = headers.length;
		for (int i = 0; i < len; i++) {
			if (headers[i] == null) {
				continue;
			}
			width[i] = countLength(headers[i]);
		}
		for (int i = 0; i < len; i++) {
			for (int j = 0; j < detectLen; j++) {
				if (data[j][i] == null) {
					continue;
				}
				int l = countLength(data[j][i]);
				width[i] = width[i] > l ? width[i] : l;
			}
		}
		for (int i = 0; i < len; i++) {
			width[i] = width[i] > MAX_COLUMN_WIDTH ? MAX_COLUMN_WIDTH
					: width[i] < MIN_COLUMN_WIDTH ? MIN_COLUMN_WIDTH : width[i];
			width[i] += 3;
		}
		return width;
	}

	private static int countLength(String s) {
		int length = 0;
		for (int i = 0; i < s.length(); i++) {
			if ((s.charAt(i) >= 0x4E00) && (s.charAt(i) <= 0x9FA5)) {
				length += 2;
			} else {
				length += 1;
			}
		}
		return length;
	}

	private boolean contains(Set<String> set, String obj) {
		for (String i : set) {
			if (i.equalsIgnoreCase(obj)) {
				return true;
			}
		}
		return false;
	}
}
