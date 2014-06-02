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
		this.workbookClass = ext EQ "xls" ? "org.apache.poi.hssf.usermodel.HSSFWorkbook" : "org.apache.poi.xssf.usermodel.XSSFWorkbook";

		// Get an input stream for the path or url file
		if (!isNull(arguments.path)) {
			// TODO: Look at using File over FileInputStream as it consumes less memory: http://poi.apache.org/spreadsheet/quick-guide.html#FileInputStrea
			this.inputStream = createObject("java", "java.io.FileInputStream").init(arguments.path);
		} else if (isDefined("arguments.url")) {
			this.inputStream = createObject("java", "java.net.URL").init(arguments.url).openStream();
		}

		if (!isNull(this.inputStream)) {
			this.workbook = createObject("java", this.workbookClass).init(this.inputStream);
		} else {
			this.workbook = createObject("java", this.workbookClass).init();
		}

		variables.cell = createObject("java", "org.apache.poi.ss.usermodel.Cell");

		return this;

	}

	function asArrays() localmode="modern" {

		arrays = [];

		for (i = 1; i < this.workbook.getNumberOfSheets(); i++) {

			sheet = this.workbook.getSheetAt(i - 1);
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

	/*function toQueries(container = "array", firstRowAsHeader = "false") localmode="modern" {


		//	How to proceed:
		//	- If firstRowAsHeader
		//		- The first row dictates which columns we parse


		//
		//	@container - The type of container to return sheet data in
		//	@firstRowAsHeader - Should the first row be treated as headers
		//

		i = 0;
		sheetsStruct = structNew("linked");
		sheetsArray = [];
		name = "";

		variables.workbook = createObject("java", "#variables.poiPackage#.xssf.usermodel.XSSFWorkbook").init(inputStream);

		for (i = 0; i < variables.workbook.getNumberOfSheets() - 1; i++)
			arrayAppend(sheetsArray, sheetToQuery(variables.workbook.getSheetAt(i)));
		}

		// DELETE? - To get the structure to maintain sheet order, we add them to the array above and then loop backwards adding to the structure as we go

		if ((arguments.container) EQ "struct") {

			for (i = 1; i < ArrayLen(sheetsArray); i++) {
				name = variables.workbook.getSheetAt(i - 1).getSheetName();
				sheetsStruct[name] = sheetsArray[i];
			}

			return sheetsStruct;

		}

		return sheetsArray;

	}*/

}