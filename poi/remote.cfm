
<cfhttp url="http://localhost:8888/Remote.cfc?method=do&returnFormat=json">

<cfdump var="#cfhttp.fileContent#">