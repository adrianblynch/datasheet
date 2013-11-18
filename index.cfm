
<!--- <cfset fis = createObject("java", "java.io.FileInputStream").init(ExpandPath("data.xls"))>
<cfset wb = createObject("java", "org.apache.poi.hssf.usermodel.HSSFWorkbook").init(fis)>
<cfset sheet = wb.getSheetAt(1)>
<cfset row = sheet.getRow(2)>
<cfset cell = row.getCell(1)>

<cfdump var="#cell.getStringCellValue()#">

<cfdump var="#cell#">

<cfabort> --->

<cfoutput>

<cfset xlsPath = ExpandPath("data.xls")>
<cfset xlsURL = "http://localhost:8888/datasheet/data.xls">
<cfset xlsPathNoExt = ExpandPath("an-xls-file")>
<cfset xlsURLNoExt = "http://localhost:8888/datasheet/an-xls-file">
<cfset xlsPathCrappy = ExpandPath("crappy-data.xls")>

<cfset xlsxPath = ExpandPath("data.xlsx")>
<cfset xlsxURL = "http://localhost:8888/datasheet/data.xlsx">
<cfset xlsxPathNoExt = ExpandPath("an-xlsx-file")>
<cfset xlsxURLNoExt = "http://localhost:8888/datasheet/an-xlsx-file">

<cfset ds = new Datasheet(path = xlsPath)>
<cfset ds = new Datasheet(url = xlsURL)>
<cfset ds = new Datasheet(path = xlsPathNoExt, ext = "xls")>
<cfset ds = new Datasheet(url = xlsURLNoExt, ext = "xls")>

<cfset ds = new Datasheet(path = xlsxPath)>
<cfset ds = new Datasheet(url = xlsxURL)>
<cfset ds = new Datasheet(path = xlsxPathNoExt, ext = "xlsx")>
<cfset ds = new Datasheet(url = xlsxURLNoExt, ext = "xlsx")>

<cfset ds = new Datasheet()>

<!---
	Read XLS file into arrays - The benefits of this method is, we can safely get data from the spreadsheet where you have holes in the data.
--->
<cfset ds = new Datasheet(path = xlsPathCrappy)>
<cfset arrays = ds.asArrays()>

<cfdump var="#arrays#">

</cfoutput>