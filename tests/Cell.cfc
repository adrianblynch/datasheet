component {

	this.CELL_TYPE_NUMERIC = 1;
	this.CELL_TYPE_STRING = 2;
	this.CELL_TYPE_BOOLEAN = 3;
	this.CELL_TYPE_BLANK = 4;
	this.CELL_TYPE_ERROR = 5;
	this.CELL_TYPE_FORMULA = 6;

	cellType = null;

	function init(cellType) {
		variables.cellType = arguments.cellType ?: null;
		return this;
	}

	function getNumericCellValue() {
		return 123;
	}

	function getStringCellValue() {
		return "abc";
	}

	function getBooleanCellValue() {
		return true;
	}

	function getCellType() {
		return cellType;
	}

}