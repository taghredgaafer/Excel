EXEC sp_configure 'show advanced options', 1
RECONFIGURE
GO
EXEC sp_configure 'ad hoc distributed queries', 1
RECONFIGURE
GO
--------------------------

EXEC sp_configure 'show advanced options', 1; 
GO 
RECONFIGURE; 
EXEC sp_configure 'ad hoc distributed Queries', 1; 
GO 
RECONFIGURE; 
EXEC sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.16.0', N'AllowInProcess', 1   
EXEC sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.16.0', N'DynamicParam', 1
-----------------------------
EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.16.0', N'AllowInProcess', 1
GO 
EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.16.0', N'DynamicParameters', 1
GO
-----------------------------
EXEC sp_configure 'show advanced options', 1; 
GO 
RECONFIGURE; 
EXEC sp_configure 'ad hoc distributed Queries', 1; 
GO 
RECONFIGURE; 
EXEC sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1   
EXEC sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParam', 1
---------------
EXEC sp_configure 'show advanced options', 1;  
GO  
RECONFIGURE;  
GO  
EXEC sp_configure 'Ad Hoc Distributed Queries', 1;  
GO  
RECONFIGURE;  
GO