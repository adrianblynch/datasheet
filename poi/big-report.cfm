
<cfquery name="sql" datasource="gri_seefusion">
	USE GRI

	DECLARE @ReportStartDate DATETIME
	DECLARE @ReportEndDate DATETIME
	DECLARE @ShowSuperUser BIT
	DECLARE @SuperUserGroupID INTEGER
	DECLARE @DeletedGroupID INTEGER
	DECLARE @DisabledGroupID INTEGER
	DECLARE @NewGroupID INTEGER
	DECLARE @NewOrgID INTEGER
	DECLARE @StartRow INTEGER
	DECLARE @EndRow INTEGER
	DECLARE @GroupID INTEGER
	DECLARE @OrgID INTEGER
	DECLARE @ClientID INTEGER
	DECLARE @ZoneID INTEGER
	DECLARE @BatchStateID INTEGER

	SET @StartRow = 1
	SET @EndRow = 0
	SET @ReportStartDate = '23-Aug-2013'
	SET @ReportEndDate = '25-Oct-2013'
	SET @ShowSuperUser = 1
	SET @SuperUserGroupID = (SELECT TOP 1 [group_id] FROM [tbl_group] WHERE [group_name] LIKE 'Super User')
	SET @DeletedGroupID = (SELECT TOP 1 [group_id] FROM [tbl_group] WHERE [group_name] LIKE '*Deleted')
	SET @DisabledGroupID = (SELECT TOP 1 [group_id] FROM [tbl_group] WHERE [group_name] LIKE '*Disabled')
	SET @NewGroupID = (SELECT [group_id] FROM [tbl_group] WHERE [group_name] LIKE '*New')
	SET @NewOrgID = (
		SELECT TOP 1 [AC].[config_value]
		FROM [CORE].[dbo].[tbl_application_config_default] AS [ACD]
		INNER JOIN [tbl_application_config] AS [AC] ON [ACD].[application_config_default_id] = [AC].[application_config_default_id]
		WHERE [config_key] LIKE 'new_org_id'
	)

	SET @BatchStateID = 2

	;WITH [ClientDownloads_CTE] AS (

		SELECT [ZB].[zipBatchId], [ZB].[processingStopDate], [Clt].[client_id], [Clt].[client_firstname], [Clt].[client_lastname], [Org].[org_id], [Org].[org_name], [Grp].[group_name], [Grp].[group_id], [ZBF].[file_id], [f].[file_name] AS [file_name]
		FROM
			[tbl_zipBatches] AS [ZB] LEFT OUTER JOIN
			[tbl_zipBatchFiles] AS [ZBF] ON [ZB].[zipBatchId] = [ZBF].[zipBatchId] LEFT OUTER JOIN
			[tbl_Client] AS [Clt] ON [ZB].[client_id] = [Clt].[client_id] LEFT OUTER JOIN
			[tbl_organisation] AS [Org] ON [Clt].[client_org_id] = [Org].[org_id] LEFT OUTER JOIN
			[tbl_group] AS [Grp] ON [Clt].[client_group_id] = [GRP].[group_id] LEFT OUTER JOIN
			[tbl_files] AS [f] ON [f].[file_id] = [ZBF].[file_id]
		WHERE ((@ShowSuperUser = 1) OR ([Grp].[group_id] <> @SuperUserGroupID))
		AND [Grp].[group_id] <> @DeletedGroupID
		AND [Grp].[group_id] <> @DisabledGroupID
		AND [Grp].[group_id] <> @NewGroupID
		AND [Org].[org_id] <> @NewOrgID
		AND [Clt].[client_id] IS NOT NULL
		AND [ZB].[zipFileBatchState] = @BatchStateID
		AND [ZB].[processingStopDate] >  @ReportStartDate
		AND [ZB].[processingStopDate] <  @ReportEndDate
	)

	, [SortedRecords_CTE] AS (
		SELECT [client_id], [client_firstname], [client_lastname], [group_id], [group_name], [org_id], [org_name], [processingStopDate], [zipBatchId], [file_name], ROW_NUMBER() OVER(ORDER BY [client_firstname] , [client_lastname] , [org_name] , [group_name] , [client_id] , [zipbatchid] , [file_name]) AS order_number, ROW_NUMBER() OVER(ORDER BY [client_firstname] DESC, [client_lastname] DESC, [org_name] DESC, [group_name] DESC, [client_id] DESC, [zipbatchid] DESC, [file_name] DESC) AS total_rows
		FROM [ClientDownloads_CTE]
		GROUP BY [client_id], [client_firstname], [client_lastname], [group_id],	[group_name], [org_id], [org_name], [processingStopDate], [zipBatchId], [file_name]
		)

	SELECT
		client_firstname + ' ' + client_lastname AS Name,
		group_name [User Group],
		org_name Organisation,
		CONVERT(VARCHAR(17), processingStopDate,13) AS [Date],
		file_name [File]
	FROM [SortedRecords_CTE]
	WHERE (@EndRow = 0 OR [order_number] BETWEEN @StartRow AND @EndRow)
	ORDER BY [order_number]
</cfquery>

<!--- Create POI --->
<cfset poi = new POI()>
<cfset workbook = poi.fromQueries([sql])>
<cfset newFilePath = expandPath(DateFormat(Now(), "ddmmyy") & "-" & TimeFormat(Now(), "HHmmss") & ".xlsx")>
<cfset fos = CreateObject("java", "java.io.FileOutputStream").init(newFilePath)>
<cfset workbook.write(fos)>
<cfset fos.close()>

<cfoutput>
	#sql.recordCount# rows writen to <a href="#ListLast(newFilePath, "/\")#">#ListLast(newFilePath, "/\")#</a>
</cfoutput>