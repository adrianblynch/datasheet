<cfoutput>

<h1>Usage</h1>

<cfset version = CreateObject("java", "org.apache.poi.Version")>

Using: #version.getProduct()# #version.getVersion()#

<cfset xlsPath = ExpandPath("sample-files/data.xls")>
<cfset xlsURL = "http://localhost:8888/datasheet/sample-files/data.xls">
<cfset xlsPathNoExt = ExpandPath("sample-files/an-xls-file")>
<cfset xlsURLNoExt = "http://localhost:8888/datasheet/sample-files/an-xls-file">
<cfset xlsPathCrappy = ExpandPath("sample-files/crappy-data.xls")>

<cfset xlsxPath = ExpandPath("sample-files/data.xlsx")>
<cfset xlsxURL = "http://localhost:8888/datasheet/sample-files/data.xlsx">
<cfset xlsxPathNoExt = ExpandPath("sample-files/an-xlsx-file")>
<cfset xlsxURLNoExt = "http://localhost:8888/datasheet/sample-files/an-xlsx-file">

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