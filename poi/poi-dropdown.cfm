
<cfset workbook = CreateObject("java", "org.apache.poi.xssf.usermodel.XSSFWorkbook").init()>
<cfset sheet = workbook.createSheet("My sheet")>

<cfset dvHelper = CreateObject("java", "org.apache.poi.xssf.usermodel.XSSFDataValidationHelper").init(sheet)>
<cfdump var="#dvHelper#" label="dvHelper">

<cfset constraintValues = "Plan A,Plan B,Plan C">
<cfset constraintValues = constraintValues.split(",")>

<cfset dvConstraint = CreateObject("java", "org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint").init(constraintValues)>

<cfset addressList = CreateObject("java", "org.apache.poi.ss.util.CellRangeAddressList").init(0, 0, 0, 0)>

<cfset validation = dvHelper.createValidation(dvConstraint, addressList)><!--- Returns an instance of org.apache.poi.xssf.usermodel.XSSFDataValidation --->
<cfset validation.setShowErrorBox(true)>
<!--- <cfset validation.setErrorStyle(DataValidation.ErrorStyle.STOP)> --->
<!--- <cfset validation.createErrorBox("Box Title", "Message Text")> --->

<cfset sheet.addValidationData(validation)>

<!--- <cfset row = sheet.createRow(0)>
<cfset cell = row.createCell(0)> --->

<cfset newFilePath = ExpandPath("dropdown.xlsx")>
<cfset fos = CreateObject("java", "java.io.FileOutputStream").init(newFilePath)>
<cfset workbook.write(fos)>
<cfset fos.close()>

<cfoutput>
	<cfset fileName = ListLast(newFilePath, "\")>
	<a href="#fileName#">#fileName#</a>
</cfoutput>

<!---
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet sheet = workbook.createSheet("Data Validation");
	XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
	XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)dvHelper.createExplicitListConstraint(new String[]{"11", "21", "31"});
	CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
	XSSFDataValidation validation = (XSSFDataValidation)dvHelper.createValidation(dvConstraint, addressList);
	validation.setShowErrorBox(true);
	sheet.addValidationData(validation);
--->