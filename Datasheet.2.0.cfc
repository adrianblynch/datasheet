<cfcomponent>
	
	<!---
		Given a file path, URL, file input stream or url input stream for an XLS or XLSX file, turn into an array or structure of queries
	--->

	<!---
		NOTES:
			- The number of columns in the returned queries can't be based solely on the number of accessible cells in the first row.
	--->

	<cffunction name="init" localmode="modern">
		
		<cfargument name="path">
		<cfargument name="url">
		<cfargument name="ext" default="xslx" hint="If a path or url doesn't have an extension, use this to force one">
		<cfargument name="stream" default="false" hint="Force the streaming version of XSSF">

		<!--- What type of file are we working with? --->

		<!--- If a path/url is given and we can work out the extension, use it --->
		<!--- If a path/url isn't given, check ext to see what sort of workbook is wanted --->

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

	<cffunction name="asQueries" localmode="modern">

		<cfargument name="firstRowAsHeaders" default="true">
		<cfargument name="ignoreEmptyRows" default="false"><!--- TODO: Implement this --->

		<cfset queries = []>

		<!--- Sheets --->
		<cfloop from="1" to="#this.workbook.getNumberOfSheets()#" index="i"><!--- NOTE: getNumberOfSheets() is 1 based --->

			<cfset sheet = this.workbook.getSheetAt(i - 1)>
			<cfset headers = []>

			<!--- NOTE: From Railo 4.2 onwards, an array can be passed to QueryNew(arrayOfColumns) instead of a list, shortening the code below --->

			<!--- Moved column logic out of the below loop to simplify things - Use first row as headers or default to Column 1, Column 2, Column N --->
			<cfif arguments.firstRowAsHeaders>
				<cfset startRow = 1>
				<cfset headers = extractCellValues(sheet.getRow(0))>
				<cfdump var="#headers#"><cfabort>
			<cfelse>
				<cfset startRow = 0>
				<cfloop from="1" to="#sheet.getRow(0).getLastCellNum()#" index="j">
					<cfset headers[j] = "Column #j#">
				</cfloop>
			</cfif>

			<cfset q = QueryNew("")>

			<cfloop array="#headers#" index="header">
				<cfset QueryAddColumn(q, header)>
			</cfloop>

			<!--- Rows --->
			<cfloop from="#startRow#" to="#sheet.getLastRowNum()#" index="j">
			
				<cfset row = sheet.getRow(j)>

				<cfif NOT IsNull(row)>

					<cfset values = extractCellValues(row)>
					<cfset QueryAddRow(q)>

					<cfloop from="1" to="#values.len()#" index="k">
						<cfset QuerySetCell(q, getColumn(q, k), values[k], j)>
					</cfloop>

				</cfif>

<cfdump var="#q#">
				
			</cfloop>

			<cfset ArrayAppend(queries, q)>

		</cfloop>

		<cfreturn queries>

	</cffunction>

	<cffunction name="getColumn" localmode="modern" hint="Get the column at a given index. If the index doesn't exist, create a default value.">

		<cfargument name="q">
		<cfargument name="index">

		<cfset headers = QueryColumnArray(arguments.q)>
		
		<!--- The column at index exists --->
		<cfif arguments.index LTE headers.len()>
			<cfset columnName = headers[arguments.index]>
		<cfelse>
			<!--- Create a new default column name --->
			<cfset columnName = "Column #arguments.index#">
			<cfset QueryAddColumn(arguments.q, columnName)>
		</cfif>

		<cfreturn columnName>

	</cffunction>

	<cffunction name="extractCellValues" localmode="modern" hint="Return all values of the given row as as array">

		<cfargument name="row">

		<cfset values = []>

		<cfloop from="0" to="#arguments.row.getLastCellNum() - 1#" index="i">

			<cfset cell = row.getCell(i, row.CREATE_NULL_AS_BLANK)>

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

				<cfset ArrayAppend(values, value)>

			</cfif>

		</cfloop>

		<cfreturn values>

	</cffunction>

</cfcomponent>