-- use yfynet
declare @table_id int
set @table_id=object_id('bempg')
dbcc showcontig(@table_id)

dbcc dbreindex('yfynet.dbo.bempg','',90)

declare @table_id int
set @table_id=object_id('bempg')
dbcc showcontig(@table_id)