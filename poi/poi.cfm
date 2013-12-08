
<!--- Create POI --->
<cfset poi = new POI()>

<!--- <cfdump var="#poi.getVersion()#" label="poi.getVersion()"><cfabort> --->

<!--- ----------------- --->
<!--- Local file system --->
<!--- ----------------- --->

<!---
	The example below reads an Excel file from disk and turns each sheet into a query, returned as an array or structure of queries.
	Those same queries are then used to generate a new Excel file which is saved to disk.
--->

<!--- A selection of files to work with --->
<cfset filePath = ExpandPath("files/events.xlsx")>
<!--- <cfset filePath = ExpandPath("files/export-import.xlsx")> --->
<!--- <cfset filePath = ExpandPath("files/planner.xlsx")> --->
<!--- <cfset filePath = ExpandPath("files/sample-events.xlsx")> --->

<!--- Read the file in --->
<cfset fileInputStream = CreateObject("java", "java.io.FileInputStream").init(filePath)>

<!--- Convert each sheet in the workbook to a query --->
<cfset queries = poi.toQueries(inputStream = fileInputStream, container = "array")>

<!--- Turn each query into a sheet --->
<cfset workbook = poi.fromQueries(queries)>

<!--- Where to save --->
<cfset newFilePath = expandPath(DateFormat(Now(), "ddmmyy") & "-" & TimeFormat(Now(), "HHmmss") & ".xlsx")>
<cfset fos = CreateObject("java", "java.io.FileOutputStream").init(newFilePath)>

<!--- Save to file --->
<cfset workbook.write(fos)>

<!--- Close the output stream --->
<cfset fos.close()>

<!--- ---------- --->
<!--- Remote URL --->
<!--- ---------- --->

<cfabort>

<cfset fileURL = "http://localhost:8888/export-import.xlsx">
<cfset fileURL = "http://localhost:8888/planner.xlsx">

<cfset urlInputStream = CreateObject("java", "java.net.URL").init(fileURL).openStream()>
<cfset sheets = poi.toQueries(urlInputStream)>
<cfdump var="#sheets#">

<!--- TODO: Get the Excel file via a redirect, does it still work? --->

<cfset fileURL = "http://localhost:8888/excel.cfm">

<cfset urlInputStream = CreateObject("java", "java.net.URL").init(fileURL).openStream()>
<cfset sheets = poi.toQueries(urlInputStream)>
<cfdump var="#sheets#">