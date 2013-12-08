
<!---
	Create an array and a structure with 1000000 indexes/keys and test how fast the look-ups are
--->

<cfset a = []>
<!--- <cfset s = {}> --->

<cfset its = 100000>

<cfset start = GetTickCount()>
<cfloop from="1" to="#its#" index="i">
	<cfset a[i] = i>
	<!--- <cfset s[i] = i> --->
</cfloop>
<cfset end = GetTickCount()>

<cfoutput>
Setup: #end - start#<br />
</cfoutput>

<cfset start = GetTickCount()>
<cfloop from="#its#" to="1" index="i" step="-1">
	<cfset found = ArrayFind(a, i)>
	<!--- <cfset found = StructKeyExists(s, i)> --->
</cfloop>
<cfset end = GetTickCount()>

<cfoutput>
Search: #end - start#<br />
</cfoutput>