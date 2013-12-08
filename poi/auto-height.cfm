
<!--- <cfset workbook = CreateObject("java", "org.apache.poi.xssf.streaming.SXSSFWorkbook").init()> --->
<!--- <cfset workbook = CreateObject("java", "org.apache.poi.hssf.usermodel.HSSFWorkbook").init()> --->

<!--- Show the bug with auto height in POI 3.8 --->

<cfset templateFile = ExpandPath("autosize-row-template.xlsx")>
<cfset fis = CreateObject("java", "java.io.FileInputStream").init(templateFile)>
<cfset workbook = CreateObject("java", "org.apache.poi.xssf.usermodel.XSSFWorkbook").init(fis)>

<cfset sheet = workbook.getSheetAt(0)>
<!--- <cfset sheet = workbook.createSheet("My sheeeeet")> --->

<cfset row = sheet.createRow(0)>

<cfset cell = row.createCell(0)>
<cfset cell.setCellValue("Line 1#Chr(10)#Line 2#Chr(10)#Line 3#Chr(10)#Line 4")>

<cfset newFilePath = ExpandPath(DateFormat(Now(), "ddmmyy") & "-" & TimeFormat(Now(), "HHmmss") & ".xlsx")>
<cfset fos = CreateObject("java", "java.io.FileOutputStream").init(newFilePath)>
<cfset workbook.write(fos)>
<cfset fos.close()>

<cfoutput>
	<cfset fileName = ListLast(newFilePath, "\")>
	<a href="#fileName#">#fileName#</a>
</cfoutput>