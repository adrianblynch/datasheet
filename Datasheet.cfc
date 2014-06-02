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
		sheets = this.workbook.iterator();
		sheetCount = 0;

		while (sheets.hasNext()) {

			sheet = sheets.next();
			rows = sheet.rowIterator();

			sheetCount++;
			arrays.append([]);
			rowCount = 0;

			while (rows.hasNext()) {

				row = rows.next();
				cells = row.cellIterator();

				rowCount++;
				arrays[sheetCount].append([]);
				cellCount = 0;

				while (cells.hasNext()) {

					cell = cells.next();

					cellCount++;
					arrays[sheetCount][rowCount].append(getCellValue(cell));

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

}