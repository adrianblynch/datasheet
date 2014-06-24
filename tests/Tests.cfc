component {

	results = {
		failures = [],
		passes = []
	};

	function setup() {

		variables.expectedData = [
			[
				["A1", "B1", "C1"],
				["A2", "B2", "C2"],
				["A3", "B3", "C3"]
			],
			[
				["A1", null, null],
				[null, "B2", null],
				[null, null, "C3"]
			],
			[
				["A1", "B1", "C1"],
				[null, null, null],
				["A3", "B3", "C3"]
			],
			[
				["", "", ""],
				["", "", ""],
				["", "", ""]
			]
		];

		return this;

	}

	function tearDown() {
		return this;
	}

	function testRegularDataAsArrays() {

		for (file in ["data.xlsx", "data.xls"]) {

			ds = getNewDS(file);
			result = ds.asArrays();

			assert(result[1][1][1] EQ expectedData[1][1][1], "Cell one in row one in sheet one in #file# is wrong", getFunctionCalledName());
			assert(result[1][3][3] EQ expectedData[1][3][3], "Cell three in row three in sheet one in #file# is wrong", getFunctionCalledName());

		}

	}

	function testMissingCellsAsArrays() {

		for (file in ["data.xlsx", "data.xls"]) {

			ds = getNewDS(file);
			result = ds.asArrays();

			assert(result[2][1][2] EQ null, "Cell two in row one in sheet 2 in #file# is not null", getFunctionCalledName());

		}

	}

	function testMissingRowsAsArrays() {

		for (file in ["data.xlsx", "data.xls"]) {

			ds = getNewDS(file);
			result = ds.asArrays();

			assert(result[3][1][1] EQ expectedData[3][1][1], "Cell one in row one in sheet 3 in #file# is wrong", getFunctionCalledName());
			assert(result[3][2][1] EQ expectedData[3][2][1], "Cell one in row two in sheet 3 in #file# is wrong", getFunctionCalledName());

		}

	}

	function testBlankCellsAsArrays() {

		for (file in ["data.xlsx"]) {

			ds = getNewDS(file);
			result = ds.asArrays();

			// assert(result[3][1][1] EQ expectedData[3][1][1], "Cell one in row one in sheet 3 in #file# is wrong", getFunctionCalledName());
			// assert(result[3][2][1] EQ expectedData[3][2][1], "Cell one in row two in sheet 3 in #file# is wrong", getFunctionCalledName());
			assert(result[4][1][1] EQ expectedData[4][1][1], "Cell one in row one in sheet 4 in #file# is wrong");
			dump(result[4]);
			assert(result[4][3][3] EQ expectedData[4][3][3], "Cell three in row three in sheet 4 in #file# is wrong");

		}

	}

	function testGetCellValue() {

		cell = new Cell();
		numericCell = new Cell(cell.CELL_TYPE_NUMERIC);
		stringCell = new Cell(cell.CELL_TYPE_STRING);
		booleanCell = new Cell(cell.CELL_TYPE_BOOLEAN);
		blankCell = new Cell(cell.CELL_TYPE_BLANK);
		errorCell = new Cell(cell.CELL_TYPE_ERROR);
		formulaCell = new Cell(cell.CELL_TYPE_FORMULA);

		ds = getNewDS();
		assert(ds.getCellValue(numericCell) EQ 123, "Numeric cell value incorrect", getFunctionCalledName());
		assert(ds.getCellValue(stringCell) EQ "abc", "String cell value incorrect", getFunctionCalledName());
		assert(ds.getCellValue(booleanCell) EQ true, "Boolean cell value incorrect", getFunctionCalledName());
		assert(ds.getCellValue(blankCell) EQ "", "Blank cell value incorrect", getFunctionCalledName());
		assert(ds.getCellValue(errorCell) EQ "", "Error cell value incorrect", getFunctionCalledName());
		assert(ds.getCellValue(formulaCell) EQ "", "Formula cell value incorrect", getFunctionCalledName());

	}

	function getNewDS(path = "data.xlsx") {
		if (!isNull(path)) {
			path = expandPath(path);
		}
		return createObject("../Datasheet").init(path = path);
	}

	function run(tests = [], abortOnFail = false) {

		variables.abortOnFail = abortOnFail;
		methodsToRun = tests;

		if (tests.len() EQ 0) {
			for (item in getMetaData(this).functions) {
				if (item.name.startsWith("test")) {
					methodsToRun.append(item.name);
				}
			}
		}

		for (method in methodsToRun) {
			setup();
			variables[method]();
		}

		return results;

	}

	function assert(condition, label = "", test = "") {
		results['#condition ? "passes" : "failures"#'].append({
			label = label,
			test = test
		});
		//checkAbortOnFail(condition);
		return condition;
	}

	function getTestName(func = "") {
		return getMetaData(this).fullName & (len(func) ? ".#func#()" : "");
	}

}