
<cfscript>

	version = createObject("java", "org.apache.poi.Version");
	echo("<p>Using: #version.getProduct()# #version.getVersion()#</p>");

	//fileName = "sample-files/crappy-data.xls";
	//fileName = "sample-files/10-11.xlsx";
	//fileName = "sample-files/10-10001.xlsx";
	//fileName = "sample-files/10-10001.xlsx";
	//fileName = "sample-files/10-64001.xlsx";

	function getSampleFileURL(fileName) localmode="modern" {
		u = cgi.request_url;
		u = listSetAt(u, listLen(u, "/"), "#fileName#", "/");
		return u;
	}

	// Local file
	/* xlsxPath = expandPath(fileName);
	ds = new Datasheet(path = xlsxPath);
	arrays = ds.asArrays();
	dump(arrays); */

	// URL
	/* xlsxURL = getSampleFileURL(fileName);
	ds = new Datasheet(url = xlsxURL);
	arrays = ds.asArrays();
	dump(arrays); */

	// Create an XLS and XLSX file using WorkbookFactory

	fileName = "sample-files/XSSFWorkbook.xlsx";
	xlsxDS = new Datasheet(path = expandPath(fileName));
	data = xlsxDS.asArrays();
	dump(data);

	fileName = "sample-files/HSSFWorkbook.xls";
	xlsDS = new Datasheet(path = expandPath(fileName));
	data = xlsDS.asArrays();
	dump(data);

</cfscript>