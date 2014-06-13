
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

	// Normal data - No cheeky stuff!
	fileName = "sample-files/1-sheet-normal-data.xlsx";
	ds = new Datasheet(path = expandPath(fileName));
	data = ds.asArrays();
	dump(data);

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

	/* fileName = "sample-files/XSSFWorkbook.xlsx";
	ds = new Datasheet(path = expandPath(fileName));
	data = ds.asArrays();
	dump(data); */

	/* fileName = "sample-files/HSSFWorkbook.xls";
	ds = new Datasheet(path = expandPath(fileName));
	data = ds.asArrays();
	dump(data); */

	// Point to extension-less files

	/* fileName = "sample-files/an-xls-file";
	ds = new Datasheet(path = expandPath(fileName));
	data = ds.asArrays();
	dump(data); */

	/* fileName = "sample-files/an-xlsx-file";
	ds = new Datasheet(path = expandPath(fileName));
	data = ds.asArrays();
	dump(data); */

	// See how crappy data is read

	/* fileName = "sample-files/crappy-data.xls";
	ds = new Datasheet(path = expandPath(fileName));
	data = ds.asArrays();
	dump(data); */

</cfscript>