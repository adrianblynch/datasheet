
<cfset tickUnit = "milli">
																														<cfset poiStart = GetTickCount(tickUnit)>
<cfset dateMask = "dd mmm yyyy">
<cfset dateMaskRegEx = REReplaceNoCase(dateMask, "\S", "\d", "all")><!--- Turn the mask into a regex --->
<cfset poi = new POI(dateMask = dateMask)>
<cfset filePath = ExpandPath("files/events.xlsx")>
<cfset fileInputStream = CreateObject("java", "java.io.FileInputStream").init(filePath)>
<cfset queries = poi.toQueries(inputStream = fileInputStream, container = "array")>
<cfset data = queries[1]>
<cfset ds = "ifg4_seefusion">
<!--- <cfset ds = "cpm_seefusion"> --->
																														<cfset poiEnd = GetTickCount(tickUnit)>
																														<cfset queriesStart = GetTickCount(tickUnit)>
<!--- Get all statuses and place in a structure for what should be a faster look-up --->
<cfquery name="allStatuses" datasource="#ds#">
	SELECT status_id id, status_description name FROM tbl_planner_status
</cfquery>

<cfset statuses = {}>

<cfloop query="allStatuses">
	<cfset statuses[Trim(allStatuses.name)] = allStatuses.id>
</cfloop>

<!--- Get all plans --->
<cfquery name="allPlans" datasource="#ds#" lazy="true">
	SELECT plan_id id, plan_name name FROM tbl_plan
</cfquery>

<cfset plans = {}>

<cfloop query="allPlans">
	<cfset plans[Trim(allPlans.name)] = allPlans.id>
</cfloop>

<!--- Get all categories --->
<cfquery name="allCats" datasource="#ds#" lazy="true">
	SELECT planner_category_id id, planner_category_name name
	FROM tbl_planner_category
</cfquery>

<cfset cats = {}>

<cfloop query="allCats">
	<cfset cats[Trim(allCats.name)] = allCats.id>
</cfloop>

<!--- Get categories in each plan --->
<cfquery name="allPlanCats" datasource="#ds#">
	SELECT pcp.planner_plan_id plan_id, pcp.planner_category_id cat_id, pc.planner_category_name name
	FROM tbl_planner_category_plan pcp
	INNER JOIN tbl_planner_category pc ON pcp.planner_category_id = pc.planner_category_id
	ORDER BY plan_id, cat_id
</cfquery>

<cfset planCats = {}>

<cfloop query="allPlanCats" group="plan_id">
	<cfset planCats[allPlanCats.plan_id] = {}>
	<cfloop>
		<cfset planCats[allPlanCats.plan_id][allPlanCats.cat_id] = allPlanCats.cat_id>
	</cfloop>
</cfloop>
																														<cfset queriesEnd = GetTickCount(tickUnit)>
<!--- Hardcoded validations till we know where the bottleneck lies --->

<cfset errors = []>
<cfset rows = []>
																														<cfset dataLoopStart = GetTickCount(tickUnit)>
<cfloop query="data">

	<cfset planID = 0>
	<cfset plan = Trim(data.Plan)>
	<cfset catID = 0>
	<cfset cat = Trim(data.Category)>
	<cfset swimlaneID = 0>
	<cfset swimlane = Trim(data.Swimlane)>
	<cfset subSwimlaneID = 0>
	<cfset subSwimlane = Trim(data["Swimlane Subcategory"][data.currentRow])>
	<cfset statusID = 0>
	<cfset status = Trim(data.Status)>

	<!--- Does status exist? --->
	<cfif NOT StructKeyExists(statuses, status)>
		<cfset ArrayAppend(errors, {"row": data.currentRow, "msg": "The status '#data.Status#' doesn't exist"})>
	</cfif>

	<!--- Does the plan exist? --->
	<cfif NOT StructKeyExists(plans, plan)>
		<cfset ArrayAppend(errors, {"row": data.currentRow, "msg": "The plan '#plan#' doesn't exist"})>
	<cfelse>
		<cfset planID = plans[plan]>
	</cfif>

	<!--- Does the category exist? --->
	<cfif NOT StructKeyExists(cats, cat)>
		<cfset ArrayAppend(errors, {"row": data.currentRow, "msg": "The category '#cat#' doesn't exist"})>
	<cfelse>
		<cfset catID = cats[cat]>
	</cfif>

	<!--- We have a plan and a category --->
	<cfif planID AND catID>

		<!--- Is the category in the plan? --->
		<cfif NOT (StructKeyExists(planCats, planID) AND StructKeyExists(planCats[planID], catID))>
			<cfset ArrayAppend(errors, {"row": data.currentRow, "msg": "The category '#cat#' is not valid in plan '#plan#'"})>
		</cfif>

	</cfif>

	<!--- Are dates valid and does the date format match what we were expecting? - Optional and maybe overly tight --->

	<cfif NOT IsDate(data.Start) OR NOT DateFormat(data.Start, dateMask) EQ data.Start>
		<cfset ArrayAppend(errors, {"row": data.currentRow, "msg": "'#data.Start#' is not a valid start date in the format '#dateMask#'"})>
	</cfif>

	<cfif NOT IsDate(data.End) OR NOT DateFormat(data.End, dateMask) EQ data.End>
		<cfset ArrayAppend(errors, {"row": data.currentRow, "msg": "'#data.End#' is not a valid start date in the format '#dateMask#'"})>
	</cfif>

	<!--- Does the start date come before the end date? --->
	<cfif IsDate(data.End) AND IsDate(data.Start) AND DateCompare(data.Start, data.End) EQ 1>
		<cfset ArrayAppend(errors, {"row": data.currentRow, "msg": "The start date, '#data.Start#', must come before the end date, '#data.End#'"})>
	</cfif>

	<cfsavecontent variable="row"><cfoutput>
		<tr>
			<td>#data.currentRow#</td>
			<td>#plan#</td>
			<td>#swimlane#</td>
			<td>#subSwimlane#</td>
			<td>#data.Title#</td>
			<td>#cat#</td>
			<td>#data.Status#</td>
			<td>#data.Start#</td>
			<td>#data.End#</td>
		</tr>
	</cfoutput></cfsavecontent>

	<cfset ArrayAppend(rows, row)>

</cfloop>
																														<cfset dataLoopEnd = GetTickCount(tickUnit)>
<!--- Begin outputting --->
<cfoutput>

	Rows in spreadsheet: #data.recordCount#<br />
	Errors found: #errors.len()#<br />
	POI: #poiEnd - poiStart# seconds (#tickUnit#)<br />
	Lookup queries: #queriesEnd - queriesStart# (#tickUnit#)<br />
	Data loop: #dataLoopEnd - dataLoopStart# (#tickUnit#)<br />
	Total: #dataLoopEnd - poiStart# seconds (#tickUnit#)

	<cfdump var="#errors#" label="errors">

	<!--- <table>
		<thead>
			<tr>
				<th>Row</td>
				<th>Plan</th>
				<th>Swimlane</th>
				<th>Swimlane Subcategory</th>
				<th>Title</th>
				<th>Category</th>
				<th>Status</th>
				<th>Start</th>
				<th>End</th>
			</tr>
		<thead>
		<tbody>
			<cfloop array="#rows#" index="row">
				#row#
			</cfloop>
		</tbody>
	</table> --->

</cfoutput>