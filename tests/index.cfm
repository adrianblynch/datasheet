<cfscript>

	totalFails = 0;
	totalPasses = 0;
	fails = [];

	tests = [
		new DatasheetTests()
	];

	for (test in tests) {
		results = test.setup().run();
		failCount = results.failures.len();
		passCount = results.passes.len();
		totalFails += failCount;
		totalPasses += passCount;
		if (failCount) {
			fails.append(results.failures);
		}
	}

	echo("Tests run: #totalFails + totalPasses#<br />");
	echo("Passes: #totalPasses#<br />");
	echo("Fails: #totalFails#<br />");

	if (totalFails GT 0) {
		header statuscode="500" statustext="#totalFails# tests failed";
		dump(fails);
	} else {
		echo("All tests passed!");
	}

</cfscript>