<cfcomponent>

	<!---
		Built and tested using POI 3.8 and .xlsx files
	--->

	<cfset VARIABLES.poiPackage = "org.apache.poi">
	<!--- TODO: Build up general constants like CELL_TYPE_NUMERIC to be queried more easily --->

	<cffunction name="init" localmode="modern">
		<cfargument name="dateMask" default="dd mmm yyyy">
		<cfset setDateMask(arguments.dateMask)>
		<cfreturn THIS>
	</cffunction>

	<cffunction name="getVersion" localmode="modern">
		<cfreturn CreateObject("java", "#VARIABLES.poiPackage#.Version").init().getVersion()>
	</cffunction>

	<cffunction name="fromQueries" localmode="modern" hint="Given one or more queries, generate a workbook">

		<cfargument name="queries" hint="Expect an array or structure of queries">

		<cfset q = "">

		<cfset VARIABLES.workbook = CreateObject("java", "#VARIABLES.poiPackage#.xssf.streaming.SXSSFWorkbook").init()>

		<cfloop collection="#queries#" item="q">
			<cfset queryToSheet(queries[q], q)>
		</cfloop>

		<cfreturn VARIABLES.workbook>

	</cffunction>

	<cffunction name="queryToSheet" localmode="modern" hint="Turn a query into a worksheet">

		<cfargument name="q">
		<cfargument name="name">

		<cfset cols = QueryColumnArray(q)>
		<cfset col = "">
		<cfset i = 0>
		<cfset sheet = VARIABLES.workbook.createSheet(name)>
		<cfset row = sheet.createRow(0)>

		<!--- Add headers --->
		<cfloop from="1" to="#ArrayLen(cols)#" index="i">
			<cfset row.createCell(i - 1).setCellValue(cols[i])>
		</cfloop>

		<!--- Add data rows --->
		<cfloop query="q">

			<cfset row = sheet.createRow(q.currentRow)>

			<cfloop from="1" to="#ArrayLen(cols)#" index="i">
				<cfset col = cols[i]>
				<cfset row.createCell(i - 1).setCellValue(q[cols[i]][q.currentRow])>
			</cfloop>

		</cfloop>

		<cfreturn sheet>

	</cffunction>

	<cffunction name="toQueries" localmode="modern" hint="Given a file input stream, assuming an .xlsx file, return an array of queries of sheets">

		<!--- TODO: Check for older .xls and use HSSFWorkbook instead --->

		<cfargument name="inputStream">
		<cfargument name="container" default="array" hint="What collection to return sheets in (struct|array)">

		<cfset i = 0>
		<cfset sheetsStruct = StructNew("linked")>
		<cfset sheetsArray = []>
		<cfset name = "">

		<cfset VARIABLES.workbook = CreateObject("java", "#VARIABLES.poiPackage#.xssf.usermodel.XSSFWorkbook").init(inputStream)>

		<cfloop from="0" to="#VARIABLES.workbook.getNumberOfSheets() - 1#" index="i">
			<cfset ArrayAppend(sheetsArray, sheetToQuery(VARIABLES.workbook.getSheetAt(i)))>
		</cfloop>

		<!--- To get the structure to maintain sheet order, we add them to the array above and then loop backwards adding to the structure as we go --->

		<cfif ARGUMENTS.container EQ "struct">

			<cfloop from="1" to="#ArrayLen(sheetsArray)#" index="i">
				<cfset name = VARIABLES.workbook.getSheetAt(i - 1).getSheetName()>
				<cfset sheetsStruct[name] = sheetsArray[i]>
			</cfloop>

			<cfreturn sheetsStruct>

		</cfif>

		<cfreturn sheetsArray>

	</cffunction>

	<cffunction name="sheetToQuery" localmode="modern">

		<cfargument name="sheet">

		<cfset q = QueryNew("")>
		<cfset headers = []>
		<cfset row = ARGUMENTS.sheet.getRow(0)>
		<cfset cellValue = "">
		<cfset data = []>
		<cfset i = 0>
		<cfset j = 0>
		<cfset DateUtil = CreateObject("java", "org.apache.poi.ss.usermodel.DateUtil")>

		<cftry>

			<cfif NOT IsNull(row)>

				<cfset noOfCells = row.getLastCellNum() - 1>
				<cfloop from="0" to="#noOfCells#" index="i">
					<cfset cell = row.getCell(i)>
					<cfset ArrayAppend(headers, cell.getStringCellValue())>
				</cfloop>

				<cfloop array="#headers#" index="header">
					<cfset QueryAddColumn(q, header)>
				</cfloop>

				<cfloop from="1" to="#ARGUMENTS.sheet.getLastRowNum()#" index="i">
					<cfset row = ARGUMENTS.sheet.getRow(i)>
					<cfset noOfCells = row.getLastCellNum() - 1>
					<cfset data = []>
					<cfloop from="0" to="#noOfCells#" index="j">
						<cfset cell = row.getCell(j)>
						<cfif NOT IsNull(cell)>
							<cfswitch expression="#cell.getCellType()#">
								<cfcase value="#cell.CELL_TYPE_NUMERIC#">
									<cfif DateUtil.isCellDateFormatted(cell)>
										<cfset cellValue = DateFormat(cell.getDateCellValue(), VARIABLES.dateMask)>
									<cfelse>
										<cfset cellValue = cell.getNumericCellValue()>
									</cfif>
								</cfcase>
								<cfcase value="#cell.CELL_TYPE_STRING#">
									<cfset cellValue = cell.getStringCellValue()>
								</cfcase>
								<cfcase value="#cell.CELL_TYPE_BOOLEAN#">
									<cfset cellValue = cell.getBooleanCellValue()>
								</cfcase>
								<cfcase value="#cell.CELL_TYPE_FORMULA#">
									<cfset cellValue = "">
								</cfcase>
								<cfcase value="#cell.CELL_TYPE_BLANK#">
									<cfset cellValue = "">
								</cfcase>
								<cfcase value="#cell.CELL_TYPE_ERROR#">
									<cfset cellValue = "">
								</cfcase>
							</cfswitch>
						<cfelse>
							<cfset cellValue = "">
						</cfif>
						<cfset ArrayAppend(data, cellValue)>
					</cfloop>
					<cfset QueryAddRow(q, data)>
				</cfloop>

			</cfif>

			<cfcatch type="any">
				<cfset q = "There was a problem reading sheet [#ARGUMENTS.sheet.getSheetName()#:#ARGUMENTS.sheet.getWorkbook().getSheetIndex(ARGUMENTS.sheet) + 1#]">
			</cfcatch>

		</cftry>

		<cfreturn q>

	</cffunction>

	<cffunction name="setDateMask" localmode="modern">
		<cfargument name="dateMask">
		<cfset variables.dateMask = arguments.dateMask>
	</cffunction>

	<cffunction name="getDateMask" localmode="modern">
		<cfreturn variables.dateMask>
	</cffunction>

</cfcomponent>