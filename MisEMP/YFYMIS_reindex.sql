use YFYMIS 
declare @table_id int
set @table_id=object_id('ysbmeord')
dbcc showcontig(@table_id)  
dbcc dbreindex('yfymis.dbo.ysbmeord','',90)

declare @table_id int
set @table_id=object_id('ysbdeord')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ysbdeord','',90) 

declare @table_id int
set @table_id=object_id('ysbmcust')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ysbmcust','',90) 

declare @table_id int
set @table_id=object_id('ysbmprod')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ysbmprod','',90) 

declare @table_id int
set @table_id=object_id('ydbdconj')
dbcc showcontig(@table_id)  
dbcc dbreindex('yfymis.dbo.ydbdconj','',90)  

declare @table_id int
set @table_id=object_id('ydbmitem')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ydbmitem','',90)  

declare @table_id int
set @table_id=object_id('ydbdptby')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ydbdptby','',90)  

declare @table_id int
set @table_id=object_id('ysbmeinv')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ysbmeinv','',90)  

declare @table_id int
set @table_id=object_id('ysbdeinv')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ysbdeinv','',90)  

declare @table_id int
set @table_id=object_id('ysbtaact')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ysbtaact','',90)  


declare @table_id int
set @table_id=object_id('ysbmdord')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ysbmdord','',90)  


declare @table_id int
set @table_id=object_id('ysbmerec')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ysbmerec','',90) 


declare @table_id int
set @table_id=object_id('ysbmcmfg')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ysbmcmfg','',90)  


declare @table_id int
set @table_id=object_id('ydbdcrpr')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ydbdcrpr','',90) 

declare @table_id int
set @table_id=object_id('YDBDEORD_TZ')
dbcc showcontig(@table_id)
 
dbcc dbreindex('yfymis.dbo.YDBDEORD_TZ','',90) 
 

declare @table_id int
set @table_id=object_id('ydbdlist')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.ydbdlist','',90)

declare @table_id int
set @table_id=object_id('YDBDEORD_TZ_Nhan')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.YDBDEORD_TZ_Nhan','',90)

declare @table_id int
set @table_id=object_id('YDBDEORD_TZ_Nhan')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.YDBDEORD_TZ_Nhan','',90) 
 
 
 declare @table_id int
set @table_id=object_id('yfypinfo')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.yfypinfo','',90) 


 declare @table_id int
set @table_id=object_id('YDBDPPGR')
dbcc showcontig(@table_id) 
dbcc dbreindex('yfymis.dbo.YDBDPPGR','',90)  

dbcc showcontig('YDBMCALE') 
dbcc dbreindex('yfymis.dbo.YDBMCALE','',90)  

dbcc showcontig('ysbtstck') 
dbcc dbreindex('yfymis.dbo.ysbtstck','',90)  

dbcc showcontig('ysbmactp') 
dbcc dbreindex('yfymis.dbo.ysbmactp','',90)  

dbcc showcontig('YDBMCONJ') 
dbcc dbreindex('yfymis.dbo.YDBMCONJ','',90)  

dbcc showcontig('YDBMPPGR') 
dbcc dbreindex('yfymis.dbo.YDBMPPGR','',90)  

dbcc showcontig('YDBMPROC') 
dbcc dbreindex('yfymis.dbo.YDBMPROC','',90)  


select*  from  cr2mis  order by FinishdateTime desc


