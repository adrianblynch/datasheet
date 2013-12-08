
<cfquery name="sql1" datasource="gri_seefusion">
USE GRI

--getDetailReportData
--"CLIENT_NAME,GROUP_NAME,ORG_NAME,PROCESSING_DATE_STRING,FILE_COUNT"
--"Name,User Group,Organisation,Date,Files"

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

				SELECT TOP 1
					[AC].[config_value]
				FROM
					[CORE].[dbo].[tbl_application_config_default] AS [ACD] INNER JOIN
					[tbl_application_config] AS [AC] ON [ACD].[application_config_default_id] = [AC].[application_config_default_id]
				WHERE
					[config_key] LIKE 'new_org_id'

			)

			SET @BatchStateID = 2


			;WITH [ClientDownloads_CTE] AS (

				SELECT
					[ZB].[zipBatchId],
					[ZB].[processingStopDate],
					[Clt].[client_id],
					[Clt].[client_firstname],
					[Clt].[client_lastname],
					[Org].[org_id],
					[Org].[org_name],
					[Grp].[group_name],
					[Grp].[group_id],
					[ZBF].[file_id]
				FROM
					[tbl_zipBatches] AS [ZB]
					LEFT JOIN [tbl_zipBatchFiles] AS [ZBF] ON [ZB].[zipBatchId] = [ZBF].[zipBatchId]
					LEFT JOIN [tbl_Client] AS [Clt] ON [ZB].[client_id] = [Clt].[client_id]
					LEFT JOIN [tbl_organisation] AS [Org] ON [Clt].[client_org_id] = [Org].[org_id]
					LEFT JOIN [tbl_group] AS [Grp] ON [Clt].[client_group_id] = [GRP].[group_id]

				WHERE
					(
						(@ShowSuperUser = 1)
						OR ([Grp].[group_id] <> @SuperUserGroupID)
					)
					AND
						[Grp].[group_id] <> @DeletedGroupID
					AND
						[Grp].[group_id] <> @DisabledGroupID
					AND
						[Grp].[group_id] <> @NewGroupID
					AND
						[Org].[org_id] <> @NewOrgID
					AND
						[Org].[show_org] = 1
					AND
						[Clt].[client_id] IS NOT NULL
					AND
						[ZB].[zipFileBatchState] = @BatchStateID
					AND
						[ZB].[processingStopDate] >  @ReportStartDate
					AND [ZB].[processingStopDate] <  @ReportEndDate



			)


			,[CountDownloads_CTE] AS (
				SELECT

					[zipBatchId],
					[processingStopDate],
					[client_id],
					[client_firstname],
					[client_lastname],
					[org_id],
					[org_name],
					[group_name],
					[group_id],
					COUNT([file_id]) AS [file_count]
				FROM
					[ClientDownloads_CTE]
				GROUP BY
					[zipBatchId],
					[processingStopDate],
					[client_id],
					[client_firstname],
					[client_lastname],
					[org_id],
					[org_name],
					[group_name],
					[group_id]
			)

			, [SortedRecords_CTE] AS (
				SELECT
					[client_id],
					[client_firstname],
					[client_lastname],
					[group_id],
					[group_name],
					[org_id],
					[org_name],
					[processingStopDate],
					[zipBatchId],
					[file_count],
					ROW_NUMBER() OVER(ORDER BY [client_firstname] , [client_lastname] , [org_name] , [group_name] , [client_id]) AS order_number,
					ROW_NUMBER() OVER(ORDER BY 	[client_firstname] DESC, [client_lastname] DESC, [org_name] DESC, [group_name] DESC, [client_id] DESC) AS total_rows
				FROM
					[CountDownloads_CTE]
				GROUP BY
					[client_id],
					[client_firstname],
					[client_lastname],
					[group_id],
					[group_name],
					[org_id],
					[org_name],
					[processingStopDate],
					[zipBatchId],
					[file_count]
			)

			SELECT
			--"CLIENT_NAME,GROUP_NAME,ORG_NAME,PROCESSING_DATE_STRING,FILE_COUNT"
			--"Name,User Group,Organisation,Date,Files"
				[client_firstname] + ' ' + [client_lastname] AS [Name],
				[group_name] [User Group],
				[org_name] [Organisation],
				CONVERT(VARCHAR(17), [processingStopDate],13) AS [Date],
				[file_count] Files
			FROM
				[SortedRecords_CTE]
			WHERE
				(@EndRow = 0
					OR
					[order_number] BETWEEN @StartRow
						AND
					@EndRow)
			ORDER BY
				[order_number]


</cfquery>
<cfquery name="sql2" datasource="gri_seefusion">

USE GRI

--getUserReportData
--"CLIENT_NAME,GROUP_NAME,ORG_NAME,DOWNLOADS,FILES"
--"Name,User Group,Organisation,Downloads,Files"

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
			DECLARE @BatchStateID INTEGER
			DECLARE @NotDownloaded BIT

			SET @StartRow = 1
			SET @EndRow = 0
			SET @ReportStartDate = '23-Aug-2013'
			SET @ReportEndDate = '25-Oct-2013'
			SET @ShowSuperUser = 1
			SET @SuperUserGroupID = (SELECT [group_id] FROM [tbl_group] WHERE [group_name] LIKE 'Super User')
			SET @DeletedGroupID = (SELECT [group_id] FROM [tbl_group] WHERE [group_name] LIKE '*Deleted')
			SET @DisabledGroupID = (SELECT [group_id] FROM [tbl_group] WHERE [group_name] LIKE '*Disabled')
			SET @NewGroupID = (SELECT [group_id] FROM [tbl_group] WHERE [group_name] LIKE '*New')
			SET @NewOrgID = (

				SELECT TOP 1
					[AC].[config_value]
				FROM
					[CORE].[dbo].[tbl_application_config_default] AS [ACD] INNER JOIN
					[tbl_application_config] AS [AC] ON [ACD].[application_config_default_id] = [AC].[application_config_default_id]
				WHERE
					[config_key] LIKE 'new_org_id'

			)

			SET @BatchStateID = 2
			SET @NotDownloaded = 0

			;WITH [ZipBatch_CTE] AS (
				SELECT [ZB].[zipBatchId],
					[ZB].[processingStopDate],
					[ZB].[client_id]
				FROM
					[tbl_zipBatches] AS [ZB]
				WHERE
					[ZB].[zipFileBatchState] = @BatchStateID
					AND (
							(
									@NotDownloaded = 0
							AND
								[ZB].[processingStopDate] >  @ReportStartDate
								AND [ZB].[processingStopDate] <  @ReportEndDate
							)
					)

			)

			, [ClientDownloads_CTE] AS (

				SELECT
					[ZB].[zipBatchId],
					[ZB].[processingStopDate],
					[Org].[org_id],
					[Org].[org_name],
					[Grp].[group_name],
					[Grp].[group_id],
					[Clt].[client_id],
					[Clt].[client_firstname],
					[Clt].[client_lastname]
				FROM
					[tbl_Client] AS [Clt] LEFT OUTER JOIN
					[tbl_organisation] AS [Org] ON [Clt].[client_org_id] = [Org].[org_id] LEFT OUTER JOIN
					[tbl_group] AS [Grp] ON [Clt].[client_group_id] = [GRP].[group_id] LEFT OUTER JOIN
					[ZipBatch_CTE] AS [ZB] ON [Clt].[client_id] = [ZB].[client_id]
				WHERE
					(
						(@ShowSuperUser = 1)
						OR ([Grp].[group_id] <> @SuperUserGroupID)
					)
					AND
						[Grp].[group_id] <> @DeletedGroupID
					AND
						[Grp].[group_id] <> @DisabledGroupID
					AND
						[Grp].[group_id] <> @NewGroupID
					AND
						[Org].[org_id] <> @NewOrgID
					AND
						[Org].[show_org] = 1
					AND
						[Clt].[client_id] IS NOT NULL


			)

			, [CountedDownloads_CTE] AS (
				SELECT
					[group_id],
					[group_name],
					[org_id],
					[org_name],
					[client_id],
					[client_firstname],
					[client_lastname],
					COUNT(DISTINCT [ClientDownloads_CTE].[zipBatchID]) AS [downloads],
					COUNT([ZBF].[zipBatchFileId]) AS [files]
				FROM
					[ClientDownloads_CTE] INNER JOIN
					[tbl_zipbatchfiles] AS [ZBF] ON [ClientDownloads_CTE].[zipBatchId] = [ZBF].[zipBatchId]
				WHERE
					@NotDownloaded = 0


				GROUP BY
					[group_id],
					[group_name],
					[org_id],
					[org_name],
					[client_id],
					[client_firstname],
					[client_lastname]

			UNION ALL

				SELECT
					[Grp].[group_id],
					[Grp].[group_name],
					[Org].[org_id],
					[Org].[org_name],
					[Clt].[client_id],
					[Clt].[client_firstname],
					[Clt].[client_lastname],
					0,
					0
				FROM
					[tbl_Client] AS [Clt] LEFT OUTER JOIN
					[tbl_organisation] AS [Org] ON [Clt].[client_org_id] = [Org].[org_id] LEFT OUTER JOIN
					[tbl_group] AS [Grp] ON [Clt].[client_group_id] = [GRP].[group_id] LEFT OUTER JOIN
					[tbl_zipBatches] AS [ZB] ON [Clt].[client_id] = [ZB].[client_id]
				WHERE
					@NotDownloaded = 1
					AND
						[ZB].[processingStopDate] IS NULL
					AND
						(
							(@ShowSuperUser = 1)
							OR ([Grp].[group_id] <> @SuperUserGroupID)
						)
					AND
						[Grp].[group_id] <> @DeletedGroupID
					AND
						[Grp].[group_id] <> @DisabledGroupID
					AND
						[Grp].[group_id] <> @NewGroupID
					AND
						[Org].[org_id] <> @NewOrgID
					AND
						[Org].[show_org] = 1
					AND
						[Clt].[client_id] IS NOT NULL


				GROUP BY
					[Org].[org_id],
					[Org].[org_name],
					[Grp].[group_name],
					[Grp].[group_id],
					[Clt].[client_id],
					[Clt].[client_firstname],
					[Clt].[client_lastname]

			)

			, [SortedRecords_CTE] AS (
				SELECT
					[group_id],
					[group_name],
					[org_id],
					[org_name],
					[client_id],
					[client_firstname],
					[client_lastname],
					[downloads],
					[files],
					ROW_NUMBER() OVER(ORDER BY [client_firstname] , [client_lastname] , [downloads] , [client_id]) AS [order_number],
					ROW_NUMBER() OVER(ORDER BY [client_firstname] DESC, [client_lastname] DESC, [downloads] DESC, [client_id] DESC) AS [total_rows]
				FROM
					[CountedDownloads_CTE]
			)

			SELECT
				--"CLIENT_NAME,GROUP_NAME,ORG_NAME,DOWNLOADS,FILES"
				--"Name,User Group,Organisation,Downloads,Files"
				[client_firstname] + ' ' + [client_lastname] AS [Name],
				[group_name] [User Group],
				[org_name] [Organisation],
				[Downloads] =
					CASE WHEN (@NotDownloaded = 0)
					THEN [downloads]
					ELSE 0
					END,

				[Files] =
					CASE WHEN (@NotDownloaded = 0)
					THEN [files]
					ELSE 0
					END
			FROM
				[SortedRecords_CTE]
			WHERE
				(@EndRow = 0
					OR
					[order_number] BETWEEN @StartRow
						AND
					@EndRow)
			ORDER BY
				[order_number]



</cfquery>
<cfquery name="sql3" datasource="gri_seefusion">


USE GRI

--getDetailReportData
--"CLIENT_NAME,GROUP_NAME,ORG_NAME,PROCESSING_DATE_STRING,FILE_COUNT"
--"Name,User Group,Organisation,Date,Files"

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

				SELECT TOP 1
					[AC].[config_value]
				FROM
					[CORE].[dbo].[tbl_application_config_default] AS [ACD] INNER JOIN
					[tbl_application_config] AS [AC] ON [ACD].[application_config_default_id] = [AC].[application_config_default_id]
				WHERE
					[config_key] LIKE 'new_org_id'

			)

			SET @BatchStateID = 2

			;WITH [ClientDownloads_CTE] AS (

				SELECT
					[ZB].[zipBatchId],
					[ZB].[processingStopDate],
					[Clt].[client_id],
					[Clt].[client_firstname],
					[Clt].[client_lastname],
					[Org].[org_id],
					[Org].[org_name],
					[Grp].[group_name],
					[Grp].[group_id],
					[ZBF].[file_id],
					[f].[file_name] AS [file_name]
				FROM
					[tbl_zipBatches] AS [ZB] LEFT OUTER JOIN
					[tbl_zipBatchFiles] AS [ZBF] ON [ZB].[zipBatchId] = [ZBF].[zipBatchId] LEFT OUTER JOIN
					[tbl_Client] AS [Clt] ON [ZB].[client_id] = [Clt].[client_id] LEFT OUTER JOIN
					[tbl_organisation] AS [Org] ON [Clt].[client_org_id] = [Org].[org_id] LEFT OUTER JOIN
					[tbl_group] AS [Grp] ON [Clt].[client_group_id] = [GRP].[group_id] LEFT OUTER JOIN
					[tbl_files] AS [f] ON [f].[file_id] = [ZBF].[file_id]
				WHERE
					(
						(@ShowSuperUser = 1)
						OR ([Grp].[group_id] <> @SuperUserGroupID)
					)
					AND
						[Grp].[group_id] <> @DeletedGroupID
					AND
						[Grp].[group_id] <> @DisabledGroupID
					AND
						[Grp].[group_id] <> @NewGroupID
					AND
						[Org].[org_id] <> @NewOrgID
					AND
						[Clt].[client_id] IS NOT NULL
					AND
						[ZB].[zipFileBatchState] = @BatchStateID
					AND
						[ZB].[processingStopDate] >  @ReportStartDate
					AND [ZB].[processingStopDate] <  @ReportEndDate


			)

			, [SortedRecords_CTE] AS (
				SELECT
					[client_id],
					[client_firstname],
					[client_lastname],
					[group_id],
					[group_name],
					[org_id],
					[org_name],
					[processingStopDate],
					[zipBatchId],
					[file_name],
					ROW_NUMBER() OVER(ORDER BY [client_firstname] , [client_lastname] , [org_name] , [group_name] , [client_id] , [zipbatchid] , [file_name]) AS order_number,
					ROW_NUMBER() OVER(ORDER BY [client_firstname] DESC, [client_lastname] DESC, [org_name] DESC, [group_name] DESC, [client_id] DESC, [zipbatchid] DESC, [file_name] DESC) AS total_rows
				FROM
					[ClientDownloads_CTE]
				GROUP BY
					[client_id],
					[client_firstname],
					[client_lastname],
					[group_id],
					[group_name],
					[org_id],
					[org_name],
					[processingStopDate],
					[zipBatchId],
					[file_name]
			)

			--"CLIENT_NAME,GROUP_NAME,ORG_NAME,PROCESSING_DATE_STRING,FILE_COUNT"
			--"Name,User Group,Organisation,Date,Files"
			SELECT
				[client_firstname] + ' ' + [client_lastname] AS [Name],
				[group_name] [User Group],
				[org_name] [Organisation],
				CONVERT(VARCHAR(17), [processingStopDate],13) AS [Date],
				[file_name] [File]
			FROM
				[SortedRecords_CTE]
			WHERE
				(@EndRow = 0
					OR
					[order_number] BETWEEN @StartRow
						AND
					@EndRow)
			ORDER BY
				[order_number]
</cfquery>

<!--- Create POI --->
<cfset poi = new POI()>
<cfset workbook = poi.fromQueries([sql1, sql2, sql3])>
<cfset newFilePath = expandPath(DateFormat(Now(), "ddmmyy") & "-" & TimeFormat(Now(), "HHmmss") & ".xlsx")>
<cfset fos = CreateObject("java", "java.io.FileOutputStream").init(newFilePath)>
<cfset workbook.write(fos)>
<cfset fos.close()>

<cfoutput>
	<a href="#ListLast(newFilePath, "/\")#">#ListLast(newFilePath, "/\")#</a>
</cfoutput>