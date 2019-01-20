# 导出List<?> 数据到excel
### 主要功能
  list<?>数据导出到excel

### 说明
  List<?> 要导出的数据
  
  String filed  json字符串， headers表头，可以多个，headersOrder 对应对象对属性
	{
    "headers":[
		"截止时间"
    ],
    "headersOrder":[
        "handleDeadline"
    ]
}

  HttpServletResponse response  
  
### 例
	// 获得需下载数据
	List<?> listObject = getDowloadList();
	XlsDownloadFilter XlsDownloadFilter = new XlsDownloadFilter(listObject, field, response);
	// 增加过滤条件
	addConvertFormat(field, XlsDownloadFilter);

	// 文件名默认为“文件-”
	XlsDownloadFilter.setFileName(fileName).start();
  
  
### 增加过滤条件
	private void addConvertFormat(String field, XlsDownloadFilter xlsDownloadFilter) {
		JSONObject obj = JSONObject.parseObject(field);
		String[] orders = jsonArray2StringArray(obj.getJSONArray("headersOrder"));
		for (String str : orders) {
			if( str.contains("Rate")) {
				xlsDownloadFilter.setConvertFormat(str, RateConvertImpl.get());
			}else if (str.contains("Date")){
				xlsDownloadFilter.setConvertFormat(str, DateFormatConvertionImp.get());
			}
		}
	}
	// json 数组转换为 string 数组
	private String[] jsonArray2StringArray(JSONArray arr) {
		String[] str = new String[arr.size()];
		for (int i = 0; i < arr.size(); i++) {
			str[i] = arr.getString(i);
		}
		return str;
	}
 
  