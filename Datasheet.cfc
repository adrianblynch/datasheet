<cfcomponent>

	<!---
		Version: 1.0
		Author: Adrian Lynch - www.adrianlynch.co.uk
	--->

	<cffunction name="init" localmode="modern">
		
		<cfargument name="path">
		<cfargument name="url">
		<cfargument name="ext" default="xslx" hint="If a path or url doesn't have an extension, use this to force one">
		<cfargument name="stream" default="false" hint="Force the streaming version of XSSF">

		<!--- Get the path or url ext --->
		<cfif IsDefined("arguments.path")>
			<cfset ext = ListLast(arguments.path, ".")>
		<cfelseif IsDefined("arguments.url")>
			<cfset ext = ListLast(arguments.url, ".")>
		<cfelse>
			<cfset ext = "xlsx">
		</cfif>

		<!--- Check the url/path ext or the one passed in to determine the workbook class to use --->
		<cfif ext EQ "xls" OR arguments.ext EQ "xls">
			<cfset this.workbookClass = "org.apache.poi.hssf.usermodel.HSSFWorkbook">
		<cfelseif ext EQ "xlsx" OR arguments.ext EQ "xlsx">
			<cfif arguments.stream>
				<cfset this.workbookClass = "org.apache.poi.xssf.streaming.SXSSFWorkbook">
			<cfelse>
				<cfset this.workbookClass = "org.apache.poi.xssf.usermodel.XSSFWorkbook">
			</cfif>
		</cfif>

		<cfif IsDefined("arguments.path")>
			<cfset this.inputStream = CreateObject("java", "java.io.FileInputStream").init(arguments.path)><!--- TODO: Look as using File over FileInputStream as it consumes less memory: http://poi.apache.org/spreadsheet/quick-guide.html#FileInputStream --->
		<cfelseif IsDefined("arguments.url")>
			<cfset this.inputStream = CreateObject("java", "java.net.URL").init(arguments.url).openStream()>
		</cfif>

		<cfif IsDefined("this.inputStream")>
			<cfset this.workbook = CreateObject("java", this.workbookClass).init(this.inputStream)>
		<cfelse>
			<cfset this.workbook = CreateObject("java", this.workbookClass).init()>
		</cfif>

		<cfset variables.cell = CreateObject("java", "org.apache.poi.ss.usermodel.Cell")>

		<cfreturn this>

	</cffunction>

	<cffunction name="asArrays" localmode="modern">
		
		<cfset arrays = []>

		<cfloop from="1" to="#this.workbook.getNumberOfSheets()#" index="i">

			<cfset sheet = this.workbook.getSheetAt(i - 1)>
			<cfset arrays[i] = []>

			<cfloop from="0" to="#sheet.getLastRowNum()#" index="j">

				<cfset row = sheet.getRow(j)>
				<cfset arrays[i][j + 1] = []>

				<cfif NOT IsNull(row)>

					<cfloop from="0" to="#row.getLastCellNum() - 1#" index="k">

						<cfset cell = row.getCell(k, row.CREATE_NULL_AS_BLANK)>
						<cfset arrays[i][j + 1][k + 1] = []>

						<cfif NOT IsNull(cell)>

							<cfset cellType = cell.getCellType()>

							<cfif cellType EQ cell.CELL_TYPE_BLANK>
								<cfset value = "">
							<cfelseif cellType EQ cell.CELL_TYPE_ERROR>
								<cfset value = "">
							<cfelseif cellType EQ cell.CELL_TYPE_FORMULA>
								<cfset value = "">
							<cfelseif cellType EQ cell.CELL_TYPE_BOOLEAN>
								<cfset value = cell.getBooleanCellValue()>
							<cfelseif cellType EQ cell.CELL_TYPE_NUMERIC>
								<cfset value = cell.getNumericCellValue()>
							<cfelseif cellType EQ cell.CELL_TYPE_STRING>
								<cfset value = cell.getStringCellValue()>
							</cfif>

							<cfset ArrayAppend(arrays[i][j + 1][k + 1], value)>

						</cfif>

					</cfloop>

				</cfif>

			</cfloop>

		</cfloop>

		<cfreturn arrays>

	</cffunction>

</cfcomponent>