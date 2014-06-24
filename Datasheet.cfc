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

		if (!isNull(arguments.path)) {
			// TODO: Look at using File over FileInputStream as it consumes less memory: http://poi.apache.org/spreadsheet/quick-guide.html#FileInputStrea
			inputStream = createObject("java", "java.io.FileInputStream").init(arguments.path);
		} else if (!isNull(arguments.url)) {
			inputStream = createObject("java", "java.net.URL").init(arguments.url).openStream();
		}

		if (!isNull(inputStream)) {
			variables.workbook = createObject("java", "org.apache.poi.ss.usermodel.WorkbookFactory").create(inputStream);
		} else {
			// FAILS: No create() method in WorkbookFactory - Which class should we create here? Maybe init() is the wrong place to determine if we're reading or writing?
			variables.workbook = createObject("java", "org.apache.poi.ss.usermodel.WorkbookFactory").create();
		}

		variables.cell = createObject("java", "org.apache.poi.ss.usermodel.Cell");
		variables.cellPolicy = "RETURN_NULL_AND_BLANK";

		return this;

	}

	function asArrays() localmode="modern" {

		arrays = [];

		i = 1;
		for (i = 1; i <= workbook.getNumberOfSheets(); i++) {

			sheet = workbook.getSheetAt(i - 1);
			arrays.append([]);
			highestCellIndex = getHighestCellIndex(sheet);

			for (j = 0; j <= sheet.getLastRowNum(); j++) {

				row = sheet.getRow(j);
				arrays[i].append([]);


				/*
					BUG: https://issues.apache.org/bugzilla/show_bug.cgi?id=30635
					Row.getLastCellNum() can report the wrong number.
					The first thought of checking for null might not work when we
					want to deal with nulls.
				*/

				// Why does this loop run once, putting a null cell in when it shouldn't - Maybe
				for (k = 0; k <= highestCellIndex; k++) {

					if (!isNull(row)) {
						cell = row.getCell(k, row[cellPolicy]);
						arrays[i][j + 1].append(getCellValue(cell));
					} else {
						arrays[i][j + 1].append(null);
					}

				}

			}

		}

		return arrays

	}

	function getHighestCellIndex(sheet) localmode="modern" {

		// To include null data when cells are skipped, get the highest cell index.

		rows = sheet.rowIterator(); // Excludes null rows - Which is OK
		highestIndex = 0;

		while (rows.hasNext()) {
			row = rows.next();
			highestIndex = max(highestIndex, row.getLastCellNum() - 1); // See comment about bug above in asArrays()
		}

		return highestIndex;

	}

	function getCellValue(cell) localmode="modern" {

		if (isNull(cell)) {
			return null;
		}

		cellType = cell.getCellType();

		if (cellType EQ cell.CELL_TYPE_NUMERIC) {
			value = cell.getNumericCellValue();
		} else if (cellType EQ cell.CELL_TYPE_STRING) {
			value = cell.getStringCellValue();
		} else if (cellType EQ cell.CELL_TYPE_BOOLEAN) {
			value = cell.getBooleanCellValue();
		} else if (cellType EQ cell.CELL_TYPE_BLANK) {
			value = "";
		} else if (cellType EQ cell.CELL_TYPE_ERROR) {
			value = "";
		} else if (cellType EQ cell.CELL_TYPE_FORMULA) {
			value = "";
		}

		return value;

	}

}