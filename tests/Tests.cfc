component {

	results = {
		failures = [],
		passes = []
	};

	function setup() {
		variables.expectedRegularData = [
			[
				["A1", "B1", "C1"],
				["A2", "B2", "C2"],
				["A3", "B3", "C3"]
			],
			[
				["A1", null, null],
				[null, "B2", null],
				[null, null, "C3"]
			]
		];
		return this;
	}

	function testRegularXLSXDataAsArrays() {

		ds = getNewDS();
		result = ds.asArrays();

		assert(result[1][1][1] EQ expectedRegularData[1][1][1], "Cell one in row one in sheet one is wrong", getFunctionCalledName());
		assert(result[1][3][3] EQ expectedRegularData[1][3][3], "Cell three in row three in sheet one is wrong", getFunctionCalledName());

	}

	function getNewDS(path = "data.xlsx") {
		return createObject("../Datasheet").init(path = expandPath(path));
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