component {

	/*
		Version: 1.0
		Author: Adrian Lynch - www.adrianlynch.co.uk
	*/

	function init(path, url) localmode="modern" {

		/*
			@path - The path to an xls(x) file
			@url - The url to an xls(x) file
		*/

		// Get the path or url ext
		ext = !isNull(arguments.path) ? listLast(arguments.path, ".") : listLast(arguments.url, ".");

		// The extension determines the type of workbook class we use
		workbookClass = ext EQ "xls" ? "org.apache.poi.hssf.usermodel.HSSFWorkbook" : "org.apache.poi.xssf.usermodel.XSSFWorkbook";

		// Get an input stream for the path or url file
		if (!isNull(arguments.path)) {
			// TODO: Look at using File over FileInputStream as it consumes less memory: http://poi.apache.org/spreadsheet/quick-guide.html#FileInputStrea
			inputStream = createObject("java", "java.io.FileInputStream").init(arguments.path);
		} else if (isDefined("arguments.url")) {
			inputStream = createObject("java", "java.net.URL").init(arguments.url).openStream();
		}

		if (!isNull(inputStream)) {
			variables.workbook = createObject("java", "org.apache.poi.ss.usermodel.WorkbookFactory").create(inputStream);
		} else {
			variables.workbook = createObject("java", workbookClass).init();
		}

		variables.cell = createObject("java", "org.apache.poi.ss.usermodel.Cell");

		return this;

	}

	function asArrays() localmode="modern" {

		arrays = [];
		sheetCount = 0;

		for (i = 1; i <= workbook.getNumberOfSheets(); i++) {

			sheet = workbook.getSheetAt(i - 1);
			arrays[i] = [];

			for (j = 0; j < sheet.getLastRowNum(); j++) {

				row = sheet.getRow(j);
				arrays[i][j + 1] = [];

				if (!isNull(row)) {

					for (k = 0; k < row.getLastCellNum() - 1; k++) {

						cell = row.getCell(k, row.CREATE_NULL_AS_BLANK);
						arrays[i][j + 1][k + 1] = [];

						if (!isNull(cell)) {
							arrayAppend(arrays[i][j + 1][k + 1], getCellValue(cell));
						}

					}

				}

			}

		}

		return arrays

	}

	function asQueries(firstRowAsHeaders = false) {

		arrays = asArrays();
		//dump(arrays);
		queries = [];

	}

	function getCellValue(cell) localmode="modern" {

		cellType = cell.getCellType();

		if (cellType EQ cell.CELL_TYPE_BLANK) {
			value = "";
		} else if (cellType EQ cell.CELL_TYPE_ERROR) {
			value = "";
		} else if (cellType EQ cell.CELL_TYPE_FORMULA) {
			value = "";
		} else if (cellType EQ cell.CELL_TYPE_BOOLEAN) {
			value = cell.getBooleanCellValue();
		} else if (cellType EQ cell.CELL_TYPE_NUMERIC) {
			value = cell.getNumericCellValue();
		} else if (cellType EQ cell.CELL_TYPE_STRING) {
			value = cell.getStringCellValue();
		}

		return value;

	}

}