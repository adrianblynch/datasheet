
<cfscript>

	version = createObject("java", "org.apache.poi.Version");
	writeOutput("<p>Using: #version.getProduct()# #version.getVersion()#</p>");

	fileName = "sample-files/10-10001.xlsx";
	//fileName = "sample-files/10-10001.xlsx";
	//fileName = "sample-files/10-64001.xlsx";

	xlsPath = expandPath(fileName);

	ds = new Datasheet(path = xlsPath);

	arrays = ds.asArrays();

	writeOutput("Read: #fileName#");

</cfscript>

<!--- Build a big XLSX file...

<cfsetting requesttimeout="9999999">

<cfset poi = new poi.POI()>

<cfset cols = []>
<cfloop from="1" to="10" index="col">
	<cfset cols.append("Header #col#")>
</cfloop>

<cfset q = QueryNew(cols.toList())>

<cfloop from="1" to="64000" index="i">
	<cfset QueryAddRow(q)>
	<cfloop from="1" to="#cols.len()#" index="j">
		<cfset QuerySetCell(q, cols[j], "#j#-#i#", i)>
	</cfloop>
</cfloop>

<cfset wb = poi.fromQueries([q])>
<cfset newFilePath = expandPath(DateFormat(Now(), "ddmmyy") & "-" & TimeFormat(Now(), "HHmmss") & ".xlsx")>
<cfset fos = CreateObject("java", "java.io.FileOutputStream").init(newFilePath)>
<cfset wb.write(fos)>
<cfset fos.close()> --->

<!--- <cfdump var="#arrays#" label="Sheet to arrays took #debug.asArraysEnd - debug.asArraysStart#ms"> --->

<!---
<!--- XLS examples --->
<cfset xlsPath = ExpandPath("sample-files/data.xls")>
<cfset xlsURL = "http://localhost:8888/datasheet/sample-files/data.xls">
<cfset xlsPathNoExt = ExpandPath("sample-files/an-xls-file")>
<cfset xlsURLNoExt = "http://localhost:8888/datasheet/sample-files/an-xls-file">
<cfset xlsPathCrappy = ExpandPath("sample-files/crappy-data.xls")>
<cfset ds = new Datasheet(path = xlsPath)>
<cfset ds = new Datasheet(url = xlsURL)>
<!--- <cfset ds = new Datasheet(path = xlsPathNoExt, ext = "xls")> --->
<!--- <cfset ds = new Datasheet(url = xlsURLNoExt, ext = "xls")> --->

<!--- XLSX examples --->
<cfset xlsxPath = ExpandPath("sample-files/data.xlsx")>
<cfset xlsxURL = "http://localhost:8888/datasheet/sample-files/data.xlsx">
<cfset xlsxPathNoExt = ExpandPath("sample-files/an-xlsx-file")>
<cfset xlsxURLNoExt = "http://localhost:8888/datasheet/sample-files/an-xlsx-file">
<cfset ds = new Datasheet(path = xlsxPath)>
<cfset ds = new Datasheet(url = xlsxURL)>
<!--- <cfset ds = new Datasheet(path = xlsxPathNoExt, ext = "xlsx")> --->
<!--- <cfset ds = new Datasheet(url = xlsxURLNoExt, ext = "xlsx")> --->

<!--- A non streamed XLSX default --->
<cfset ds = new Datasheet()>

<!--- Read XLS file into arrays - The benefits of this method is, we can safely get data from the spreadsheet where you have holes in the data. --->
<cfset ds = new Datasheet(path = xlsPathCrappy)>
<cfset arrays = ds.asArrays()>

<cfdump var="#arrays#" expand="false">

<!--- To come --->
<!--- Read XLS into queries --->
<!--- <cfset ds = new Datasheet(path = xlsPathCrappy)>
<cfset queries = ds.asQueries()> ---> --->