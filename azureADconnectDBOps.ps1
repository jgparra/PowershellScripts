<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.3.131
	 Created on:   	2/17/2017 1:15 AM
	 Created by:   	Gregorio Parra - (gregorio.parra@microsoft.com)
	 				Thank you to Hernan Bohl (hbohl@microsoft.com)
					and  Daniel Valero  (Daniel.Valero@microsoft.com) 
					
	 Organization: 	Microsoft
	 Filename:     	azureADconnectDBOps.ps1
	===========================================================================
	.DESCRIPTION
		this Tool is not supported by Microsoft!!
		please, all comments ans suggestions here - 
		
		this tool is created to see information from Azure AD Sync Database
		this information will be useful to check space issues for database
		or to see tables consuming space inside database
		
		versions
		V 1.0  2/17/2017 - initial version
#>

#region transactSpaceused
<#
this SQL command is used to check database version, O.S. version and to see 
space consumed by database
#>
$transactSpaceused=@"
select @@version 
go
use ADSync 
go 
exec SP_SPACEUSED 
go 
"@

#endregion transactSpaceused

#region transactObjectCount
<#
this SQL query will show the number of objects
#>
$transactObjectCount= @"
SELECT object_type,count (*) as cuenta
  FROM [ADSync].[dbo].[mms_metaverse]
  group by object_type
SELECT count (*)
  FROM [ADSync].[dbo].[mms_connectorspace]
go
"@
#endregion transactObjectCount

#region transactErrorLog

$transactErrorLog=@"
exec xp_readerrorlog 1
go
"@
#endregion transactErrorLog

#region transactFragmentation

$transactFragmentation=@"
use ADSync 
go 
select O.name as tabla, I.name as indice, F.index_type_desc, F.page_count, F.avg_fragmentation_in_percent  
from sys.dm_db_index_physical_stats(db_id('ADSync'),null,null,null,'LIMITED') F
inner join sys.indexes I on F.index_id = I.index_id  and F.object_id = I.object_id
inner join sys.objects O on O.object_id = I.object_id
where avg_fragmentation_in_percent > 10
and page_count > 10
go
"@
#endregion transactFragmentation

#region transacttablespace

$transacttablespace =@"
use ADSync 
go 
SELECT 
    t.NAME AS TableName,
    s.Name AS SchemaName,
    p.rows AS RowCounts,
    SUM(a.total_pages) * 8 AS TotalSpaceKB, 
    SUM(a.used_pages) * 8 AS UsedSpaceKB, 
    (SUM(a.total_pages) - SUM(a.used_pages)) * 8 AS UnusedSpaceKB
FROM 
    sys.tables t
INNER JOIN      
    sys.indexes i ON t.OBJECT_ID = i.object_id
INNER JOIN 
    sys.partitions p ON i.object_id = p.OBJECT_ID AND i.index_id = p.index_id
INNER JOIN 
    sys.allocation_units a ON p.partition_id = a.container_id
LEFT OUTER JOIN 
    sys.schemas s ON t.schema_id = s.schema_id
WHERE 
    t.NAME NOT LIKE 'dt%' 
    AND t.is_ms_shipped = 0
    AND i.OBJECT_ID > 255 
GROUP BY 
    t.Name, s.Name, p.Rows
ORDER BY 
    UsedSpaceKB desc
go
"@
#endregion transacttablespace

function execSQLCMD {
	param ($SQLAction)
	
	$instancia = (get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server Local DB\Shared Instances\ADSync" -Name InstanceName).instancename
	Write-Host "instancia: $instancia" -ForegroundColor Cyan
	
#	$comando = "sqlcmd -S np:\\.\pipe\$instancia\tsql\query"
	$comando = "sqlcmd -S np:\\.\pipe\$instancia\tsql\query -Q `"$SQLAction`" "
	
	Write-Host "ejecutando: $comando"
	Invoke-Expression $comando	
}

#region menu
$menu = {
	Clear-Host
	write-host "	******************************************************************" -foregroundcolor cyan
	write-host "	Dirsync Database verification tool Version 1.0" -foregroundcolor cyan
	write-host "	******************************************************************" -foregroundcolor cyan
	write-host "	"
	write-host "	Please select an option from the list below." -foregroundcolor yellow
	write-host "	"
	write-host "	Dirsync" -foregroundcolor yellow
	write-host "	1) get space used " -foregroundcolor white
	write-host "	2) get object count" -foregroundcolor white
	write-host "	3) get error log " -foregroundcolor white
	write-host "	4) get db fragmentation" -foregroundcolor white
	write-host "	5) get table space used" -foregroundcolor white
	write-host "		"
	write-host "	98) Restart the Server" -foregroundcolor red
	write-host "	99) Exit" -foregroundcolor cyan
	write-host "	"
	write-host "	Select an option.. [1-99]?  " -foregroundcolor white -nonewline
	
	
}
#endregion menu

do {
	invoke-command -scriptblock $Menu
	$condition = Read-Host
	switch ($condition) {
		1 {
			# get space used
			execSQLCMD $transactSpaceused
		}
		2 {
			#get object count
			execSQLCMD $transactObjectCount
		}
		3 {
			#get error log
			execSQLCMD $transactErrorLog
		}
		4 {
			#get database fragmentation
			execSQLCMD $transactFragmentation
		}
		5 {
			#get database fragmentation
			execSQLCMD $transacttablespace
		}
		77 {
			#restart server
			restart-computer -computername localhost -force
		}
		99 {
			#exit
			Pop-Location
			Write-Host "Exiting..."
		}
		default {
			#<code>
		}
	}
	Write-Host "press Enter to continue" -ForegroundColor Green
	Read-Host
}
while ($condition -ne 99)


