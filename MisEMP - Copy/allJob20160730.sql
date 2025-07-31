USE [msdb]
GO
/****** 物件:  Job [EMP_更新員工特休假]    指令碼日期: 07/30/2016 02:10:42 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:42 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'EMP_更新員工特休假', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:42 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare  @dat1  as  varchar(10)  
declare  @dat2  as  datetime  

--- 每年于3月份結算員工未休之年假  (未休假者於3月份薪資中 給付假日津貼 ) 
--- EX: 08年計算 07 年所剩下的年假 ,  由 2006/4/1 開始計算  
if  right( convert(char(10), getdate(),111) ,5 ) <= ''04/01'' 
	begin 
	set  @dat1 = ltrim(rtrim(str(year(getdate())-1)) ) +''/4/1''   
	end   
print @dat1 

 
update empfile   set  tx =   
case when   floor(  DATEDIFF( m,   b.calctx   , isnull(outdat,getdate()) )    )    <=0 then 0 
else  
	 floor(  DATEDIFF(m,    b.calctx   , isnull(outdat,getdate()) )    )   
	  +   ( case when  floor( DATEDIFF(m,  b.calctx  , isnull(outdat,getdate()) ) /30.00 )  >=60  then floor ( cast ( year(  isnull(outdat,getdate()) )  - year(indat)   as int  )  /5 )   else  0 end  ) 
end  
from empfile a , (select  empid , calctx from view_empfile ) b 
where  a.empid = b.empid', 
		@database_name=N'YFYNET', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=16, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=1, 
		@active_start_date=20080310, 
		@active_end_date=99991231, 
		@active_start_time=50000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [EMP_員工薪資備份]    指令碼日期: 07/30/2016 02:10:43 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:43 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'EMP_員工薪資備份', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:43 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare @dat   as datetime
declare @BKym as varchar(6) 
set  @dat  = dateadd( d,-2 ,  getdate() )  
set @BKym = ltrim(rtrim(str(year(@dat))))+right( ''00'' + ltrim(rtrim(str(month(@dat)))) , 2 ) 
------------------------------------------------------------------------
delete  empdsalary_bak where  yymm=@BKym  
insert into  empdsalary_bak  
select   getdate(), ''SYS'' ,  *    from empdsalary  where     yymm=@BKYM   

------------------------------------------------------------------------ 
--delete  VYFYMYJXX  where  yymm=@bkym  
--insert into  VYFYMYJXX 
--select    * , getdate(), ''SYS''  from VYFYMYJX where  yymm=@bkym  

------------------------------------------------------------------------ 
--delete  EMPBHGTX  where  yymm=@bkym  
--insert into EMPBHGTX  
--select    * , getdate(), ''SYS''  from EMPBHGT where  yymm=@bkym  

---------------------------------------------------------------------
--delete  EMPWORKX  where  yymm=@bkym  
--insert into EMPWORKX  
--select    * , getdate(), ''SYS''  from EMPWORK where  yymm=@bkym   
', 
		@database_name=N'YFYNET', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=16, 
		@freq_interval=2, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=1, 
		@active_start_date=20070501, 
		@active_end_date=99991231, 
		@active_start_time=180000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [EMP_新增每月日期(薪資系統)]    指令碼日期: 07/30/2016 02:10:43 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:43 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'EMP_新增每月日期(薪資系統)', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:43 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare @days  as int   
declare @days1  as int   
declare @days2  as int   
declare  @cDatestr as datetime 
declare  @cDatestr1 as datetime 
declare  @cDatestr2 as datetime 
set   @cDatestr= getdate()
set   @cDatestr1=  getdate()+30
set   @cDatestr2= getdate()+60 
--print   @cDatestr 
--print   @cDatestr1 
--print   @cDatestr2 
set @days=DAY(@cDatestr+(32-DAY(@cDatestr))-DAY(@cDatestr+(32-DAY(@cDatestr))))   
set @days1=DAY(@cDatestr1+(32-DAY(@cDatestr1))-DAY(@cDatestr1+(32-DAY(@cDatestr1))))   
set @days2=DAY(@cDatestr2+(32-DAY(@cDatestr2))-DAY(@cDatestr2+(32-DAY(@cDatestr2))))   
--print @days 
--print @days1 
--print @days2  
declare @i as int 
declare @i1 as int 
declare @i2 as int 
declare @Datestr   as  varchar(10)  
declare @Datestr1   as  varchar(10)  
declare @Datestr2   as  varchar(10)  
declare @NowMonth   as varchar(6) 
declare @NowMonth1   as varchar(6) 
declare @NowMonth2   as varchar(6) 

set   @NowMonth =  ltrim(rtrim(str(year(getdate()))))+ right(''00''+ltrim(rtrim(str(month(getdate())))) ,2) 
set   @NowMonth1 =  ltrim(rtrim(str(year(@cDatestr1))))+ right(''00''+ltrim(rtrim(str(month(@cDatestr1)))) ,2)    --right(''00''+convert(varchar(6),@cDatestr1,112) ,2)
set   @NowMonth2 =   ltrim(rtrim(str(year(@cDatestr2))))+ right(''00''+ltrim(rtrim(str(month(@cDatestr2)))) ,2)    --right(''00''+convert(varchar(6),@cDatestr2,112) ,2)

set @i=1   
set @i1=1   
set @i2=1   

declare  @sts  as varchar(5)  

while   @i  <= @days  
   begin 
	set @Datestr = ltrim(rtrim(str(left(@NowMonth,4))))+''/''+right(@NowMonth,2)+''/''+right(''00''+ltrim(rtrim( str(@i))),2)  
	print @Datestr 
	 set @i=@i+1    
	if not exists ( select * from  YDBMCALE  where convert(char(10),dat,111)=convert(char(10),@Datestr,111) and convert(char(6), dat, 112) = @NowMonth  ) 
		begin 
		if  DATEPART (weekday, @Datestr )  = 1 
			set  @sts  =''H2'' 
		else
			set  @sts  =''H1'' 
		insert into  YDBMCALE  (dat, STATUS) values (  @Datestr, @sts ) 
		end 
    end  
while   @i1  <= @days1  
   begin 
	set @Datestr1 = ltrim(rtrim(str(left(@NowMonth1,4))))+''/''+right(@NowMonth1,2)+''/''+right(''00''+ltrim(rtrim( str(@i1))),2)  
	print @Datestr1 
	 set @i1=@i1+1    
	if not exists ( select * from  YDBMCALE  where convert(char(10),dat,111)=convert(char(10),@Datestr1,111) and convert(char(6), dat, 112) = @NowMonth1  ) 
		begin 
		if  DATEPART (weekday, @Datestr )  = 1 
			set  @sts  =''H2'' 
		else
			set  @sts  =''H1''  
		insert into  YDBMCALE  (dat, STATUS) values (  @Datestr1, @sts ) 
		end 
    end   
while   @i2  <= @days2  
   begin 
	set @Datestr2 = ltrim(rtrim(str(left(@NowMonth2,4))))+''/''+right(@NowMonth2,2)+''/''+right(''00''+ltrim(rtrim( str(@i2))),2)  
	print @Datestr2
	 set @i2=@i2+1    
	if not exists ( select * from  YDBMCALE  where convert(char(10),dat,111)=convert(char(10),@Datestr2,111) and convert(char(6), dat, 112) = @NowMonth2  ) 
		begin 
		if  DATEPART (weekday, @Datestr )  = 1 
			set  @sts  =''H2'' 
		else
			set  @sts  =''H1'' 	
		insert into  YDBMCALE  (dat, STATUS) values (  @Datestr2, @sts ) 
		end 
    end 
', 
		@database_name=N'yfynet', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=16, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=1, 
		@active_start_date=20080103, 
		@active_end_date=99991231, 
		@active_start_time=1000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [EMP_薪資職務單位更新]    指令碼日期: 07/30/2016 02:10:44 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:44 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'EMP_薪資職務單位更新', 
		@enabled=1, 
		@notify_level_eventlog=0, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [EMP_薪資職務單位更新]    指令碼日期: 07/30/2016 02:10:44 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'EMP_薪資職務單位更新', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare  @Lym as varchar(6)
declare @Nym as varchar(6) 
set @Lym =  ltrim(rtrim(str(year(getdate()-1))))+ right(''00''+ltrim(rtrim(str(month(getdate()-1)))),2)
set @Nym = ltrim(rtrim(str(year(getdate()))))+ right(''00''+ltrim(rtrim(str(month(getdate())))),2)
--print @Nym
--print @Lym
--單位
insert into bempg ( yymm,  whsno, empid,  country, groupid, zuno, shift, memo , mdtm, muser ) 
select  @Nym , a.whsno  ,  a.empid, c.country, a.groupid ,a.zuno , a.shift   , ''''  ,getdate(), ''SYS''   from  
( select * from bempg where  yymm= @Lym ) a 
join empfile c on c.empid = a.empid   
where (  isnull(outdat,'''')='''' or  convert(char(10), isnull(outdat,''''),111) >= convert(char(10),getdate(),111)   ) 
and a.empid not in ( select  empid from bempg where  yymm=@Nym )
 
--職務
insert into bempj ( yymm, whsno, country, empid, job, memo, mdtm , muser ) 
select  @Nym, a.whsno, a.country, a.empid,  a.job, '''', getdate(), ''SYS''  from 
( select * from bempj where  yymm= @Lym ) a 
join empfile b on b.empid = a.empid 
where (  isnull(outdat,'''')='''' or  convert(char(10), isnull(outdat,''''),111) >= convert(char(10),getdate(),111)   ) 
and a.empid not in ( select  empid from bempj where  yymm=@Nym  )', 
		@database_name=N'yfynet', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'EMP_薪資職務單位更新', 
		@enabled=1, 
		@freq_type=32, 
		@freq_interval=8, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=1, 
		@freq_recurrence_factor=1, 
		@active_start_date=20121005, 
		@active_end_date=99991231, 
		@active_start_time=40000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_入庫明細備份]    指令碼日期: 07/30/2016 02:10:44 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:44 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_入庫明細備份', 
		@enabled=1, 
		@notify_level_eventlog=0, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:44 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare @yymm  as varchar(6) 
set @yymm   = convert(char(6),getdate()-1,112)

exec sp_入庫明細備份 @yymm     ', 
		@database_name=N'yfymis', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch1', 
		@enabled=1, 
		@freq_type=32, 
		@freq_interval=8, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=1, 
		@freq_recurrence_factor=1, 
		@active_start_date=20130618, 
		@active_end_date=99991231, 
		@active_start_time=30000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_已收托製資料回傳各廠]    指令碼日期: 07/30/2016 02:10:45 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:45 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_已收托製資料回傳各廠', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:45 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'INSERT INTO [172.22.169.21].[YFYMIS].[dbo].[YSBDEORD_TZ] 
([FrW1], [custid], [parentsc#], [sc#], [comp#], [status], [po#], [lot#], [ordSqty], [ordSadd], [ratio], [model#], [part_Code], 
[descp], [keyinby], [keyindate], [mdtm] , utp  ) 
select   ''LA'' , a.custid, b.parentsc# , a.sc#, a.comp#, a.sts , b.[p/o#], b.[lot#] ,  
a.delqty, a.addqtys,  case when left(a.sc#,1)=''A'' then c.mdout else a.ratio end as ratio , b.model#, b.part_code, b.[description] , 
b.keyinby ,  b.keyindate,getdate()  , b.unitprice 
from 
(select  myordno  from ydbdeord_tz_nhan f where  frmW1=''DN'' and isnull(tzflag,'''')=''*'' ) zx 
left join  ( Select parentsc#, sc#, comp#, custid, model#, part_code, [description], [p/o#], lot# , keyindate, keyinby , unitprice  from ysbmeord   with(nolock)   ) b  on b.sc# =zx.myordno
join ( 
	select  custid,parentsc#,  sc#, comP# , status , delqty , additionalallowance as addqtys , ratio ,
	case when status=''Cancel'' then ''D'' else left(status,1) end as sts  ,
	sc#+case when status=''Cancel'' then ''D'' else left(status,1) end as T1 
	from ysbdeord    with(nolock)  where  item#=''1''  
 ) a  on a.sc# = b.sc# 
left join (select parentcomp#, comp#, ratio, [out] as mdout  from ysbmprod   with(nolock)  ) c on c.comp# = a.comp#  
left join ( select sc#+left(status,1) as T1 from [172.22.169.21].[YFYMIS].[dbo].[YSBDEORD_TZ] ) x on x.t1 = a.t1 
where isnull(x.t1,'''')=''''
 

INSERT INTO [172.22.171.33].[YFYMIS].[dbo].[YSBDEORD_TZ] 
([FrW1], [custid], [parentsc#], [sc#], [comp#], [status], [po#], [lot#], [ordSqty], [ordSadd], [ratio], [model#], [part_Code], 
[descp], [keyinby], [keyindate], [mdtm] , utp ) 
select   ''LA'' , a.custid, b.parentsc# , a.sc#, a.comp#, a.sts , b.[p/o#], b.[lot#] ,  
a.delqty, a.addqtys,  case when left(a.sc#,1)=''A'' then c.mdout else a.ratio end as ratio , b.model#, b.part_code, b.[description] , 
b.keyinby , b.keyindate, getdate()  , unitprice 
from 
(select  myordno  from ydbdeord_tz_nhan f where  frmW1=''BC'' and isnull(tzflag,'''')=''*'' ) zx 
left join  ( Select parentsc#, sc#, comp#, custid, model#, part_code, [description], [p/o#], lot# , keyindate, keyinby , unitprice  from ysbmeord    with(nolock)    ) b  on b.sc# =zx.myordno
join ( 
	select  custid,parentsc#,  sc#, comP# , status , delqty , additionalallowance as addqtys , ratio ,
	case when status=''Cancel'' then ''D'' else left(status,1) end as sts  ,
	sc#+case when status=''Cancel'' then ''D'' else left(status,1) end as T1 
	from ysbdeord   with(nolock)  where  item#=''1''  
 ) a  on a.sc# = b.sc# 
left join (select parentcomp#, comp#, ratio, [out] as mdout  from ysbmprod   with(nolock)  ) c on c.comp# = a.comp#  
left join ( select sc#+left(status,1) as T1 from [172.22.171.33].[YFYMIS].[dbo].[YSBDEORD_TZ] ) x on x.t1 = a.t1 
where isnull(x.t1,'''')='''' 


delete  ysbdeord_tz where  aid not in ( 
select b.aid  from 
(select max(status) status, sc#  from  ysbdeord_tz  group by sc# )a
left join ( select aid, status, sc# from ysbdeord_tz ) b on b.sc# = a.sc# and b.status= a.status
)  
 
 
 
', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=8, 
		@freq_subday_interval=5, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20091021, 
		@active_end_date=99991231, 
		@active_start_time=40000, 
		@active_end_time=215959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_平政同奈廠事故資料]    指令碼日期: 07/30/2016 02:10:45 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:45 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_平政同奈廠事故資料', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:45 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare @yymm as varchar(6) 
set @yymm = convert(char(6),getdate(),112)
--print @yymm 


--  同奈    
delete   yfydsgkm_dn  where (sgym =@yymm  or  yymm=@yymm )     

insert into yfydsgkm_dn 
select * from  [172.22.169.21].[yfymis].dbo.yfydsgkm where  (sgym =@yymm  or  yymm=@yymm )   


delete  yfymsuco_dn where 
autoid+sgno+sgym in ( 
select autoid+sgno+sgym  from  [172.22.169.21].[yfymis].dbo.yfydsgkm where  yymm=@yymm  or  sgym=@yymm  
)    

insert into  yfymsuco_dn  
select * from  [172.22.169.21].[yfymis].dbo.yfymsuco where  
autoid+sgno+sgym in ( 
select   autoid+sgno+sgym  from  [172.22.169.21].[yfymis].dbo.yfydsgkm where  yymm=@yymm or  sgym=@yymm 
)  

--  平政 

delete   yfydsgkm_bc  where  (sgym =@yymm  or  yymm=@yymm )    

insert into yfydsgkm_bc 
select * from  [172.22.171.33].[yfymis].dbo.yfydsgkm where  (sgym =@yymm  or  yymm=@yymm )   


delete  yfymsuco_bc where 
autoid+sgno+sgym in ( 
select autoid+sgno+sgym  from  [172.22.171.33].[yfymis].dbo.yfydsgkm where   (sgym =@yymm  or  yymm=@yymm )   
)    

insert into  yfymsuco_bc  
select * from  [172.22.171.33].[yfymis].dbo.yfymsuco where  
autoid+sgno+sgym in ( 
select   autoid+sgno+sgym  from  [172.22.171.33].[yfymis].dbo.yfydsgkm where  (sgym =@yymm  or  yymm=@yymm )   
)   
', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20090429, 
		@active_end_date=99991231, 
		@active_start_time=220000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_更新平板送貨通知]    指令碼日期: 07/30/2016 02:10:46 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:46 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_更新平板送貨通知', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:46 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'update  ysbmeord set  status=''InDel''    where  left(sc#,1)=''A''
and isnull(status,'''')='''' 
  ', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=8, 
		@freq_subday_interval=1, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20070413, 
		@active_end_date=99991231, 
		@active_start_time=0, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_更新受訂總額]    指令碼日期: 07/30/2016 02:10:46 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:46 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_更新受訂總額', 
		@enabled=1, 
		@notify_level_eventlog=0, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [sch01]    指令碼日期: 07/30/2016 02:10:47 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'sch01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare  @yymm as varchar(6) 
set @yymm = convert(varchar(6),getdate(),112)

update  ysbdclmt set  finmonthtodatesale = isnull(toamt,0)
from ysbdclmt a , 
( 
	select case when left(sc#,1)=''A'' then ''SHB'' else ''CTN'' end as comptype , custid , salesman, sum(po_qty*unitprice*isnull(exrt,1) ) as toamt  from 
	ysbmeord  a 
	left join ( select * from ysbmexrt ) b on b.code = a.dm and b.yyyymm = convert(Char(6),keyindate ,112) 
	join (select comp#, delinset from ysbmprod ) c on c.comp# = a.comp# 
	where 
	c.delinset<>''Typebs'' and  scstatus<>''cancel'' and  convert(Char(6),keyindate,112)=@yymm  and left(sc#,1)<>''E''    
	group by custid ,salesman , case when left(sc#,1)=''A'' then ''SHB'' else ''CTN'' end 
) b where  a.custid = b.custid and  a.papertype= b.comptype and a.salesman = b.salesman 
and round(isnull(a.finmonthtodatesale,0),0) <> round(isnull(b.toamt,0),0)', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'step01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=8, 
		@freq_subday_interval=1, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20080528, 
		@active_end_date=99991231, 
		@active_start_time=0, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_更新狀態]    指令碼日期: 07/30/2016 02:10:47 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:47 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_更新狀態', 
		@enabled=0, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [sch01]    指令碼日期: 07/30/2016 02:10:47 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'sch01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'update   ysbdeord set status=''*''    
where  status<>''cancel'' and   status<>''*''  and len(item#)=1 and sc# in ( 
select  sc#  from  ysbmdord where  isnull(do#,'''')<>'''' and  status=''1'' and  isnull(goodsstatus,'''')<>''cancel''   
)   

update   ysbdeord set status=''''    
where  status<>''cancel'' and   status=''*''  and len(item#)=1 and sc# in ( 
select  sc#  from  ysbmdord where  isnull(do#,'''')<>'''' and  status=''1'' and  isnull(goodsstatus,'''')=''cancel''   
)   

update  ysbdeord set status=''Z''  
where    status<>''cancel'' and  status<>''*''  and status<>''Z''  and sc#+item#     in (  
select  sc#+item as sc2    from  ydbdlist      group by   sc#+item   having sum(case when status=''D1'' then -1*actqty else actqty end ) > 0  
)     

update   ysbdeord set status=''*''    
where  status<>''cancel'' and   status<>''*''  and len(item#)=1 and sc# in ( 
select  sc#  from  ysbmdord where  isnull(do#,'''')<>'''' and  status=''1'' and  isnull(goodsstatus,'''')<>''cancel''   
)    

update  ysbdeord set status=''1''  
 where  status<>''cancel'' and status<>''*'' and status<>''Z''  and status<>''1'' and sc#+item# in ( 
select sc#+item# from ydbmitem where  isnull(finishdatetime,'''')<>'''' 
)   

update  ysbdeord set status=''B''  
where  status<>''cancel'' and status<>''*'' and status<>''Z''  and status<>''1''  and status<>''B'' and sc#+item# in ( 
select sc#+item# from ydbmitem where    isnull(seq#,'''')<>''''
)   


update  ysbdeord set status=''A''  
where  status<>''cancel'' and status<>''*'' and status<>''Z''  and status<>''1''  and status=''B'' and sc#+item# not  in ( 
select sc#+item# from ydbmitem where    isnull(seq#,'''')<>''''
)', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'step01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=4, 
		@freq_subday_interval=10, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20071113, 
		@active_end_date=99991231, 
		@active_start_time=0, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_更新產品重量面積]    指令碼日期: 07/30/2016 02:10:47 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:47 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_更新產品重量面積', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:48 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'--------加總計算產品重量、面積  (2006/09/22)   
insert into  yfypinfo  
select a.* from ( 
	select  custid, PARENTCOMP#,  sum(calcWeight)  as unitweight,   sum( case when left(parentcomp#,1)=''S'' then area  else area*ratio*firstorder end ) as area, sum(new_kg)as totWeight , count(*) as prodLine  , getdate() mdtm   from ( 
		SELECT a.custid, A.PARENTCOMP#,  A.COMP#,  A.BOARDQULITY, A.FLUTE,  A.[SQM/PIECE] AREA,  A.RATIO, A.FIRSTORDER, 
		round( (e.unitweight)/1000 , 3 )  as calcWeight , 
		case when left(PARENTCOMP#,1)=''S'' then 
			round(  (e.unitweight/1000) , 3)   *  round( A.[sqm/piece],5) 
		else
			round(  (e.unitweight/1000) , 3)   *  round( A.[sqm/piece],5)   * A.ratio * A.firstorder    
		end as  PAPERWeight,   
		case when left(parentcomp#,1)=''S'' then 
			round( (e.unitweight)/1000 , 3 ) * A.[sqm/piece] 
		else
			round( (e.unitweight)/1000 , 3 ) * A.[sqm/piece]  * A.ratio * A.firstorder    
		end as   new_kg 
		FROM 
		( SELECT *  FROM YSBMPROD WHERE DELINSET<>''TYPED''    ) A 
		 inner join  (  select * from ysbmtqty ) e on  e.boardquality = A.boardqulity and e.papertype = A.comptype and e.flute = A.flute   
	) z group by custid,  PARENTCOMP#    
) a 
left join ( select *  from yfypinfo  ) b on b.parentcomp# = a.parentcomp# 
where  isnull(b.parentcomp#,'''')='''' 

--relation  CheckCost_CTN  , CheckCost_SHB  
--step2 ---------------------------------------------------------------------------
select  parentcomp# ,  sum([sqm/piece]*case when left(comp#,1)=''S'' then 1 else ratio*firstorder end ) as area  into #tmparea   from ysbmprod  where  delinset<>''typed''  group  by parentcomp#   

update   yfypinfo set  prodarea = b.area  
from   yfypinfo a, #tmparea  b 
where   a.parentcomp# = b.parentcomp# 
and  round(a.prodarea,3) <>  round(b.area,3)   
and isnull(b.parentcomp#,'''')<>'''' 
 
', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=30, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20070327, 
		@active_end_date=99991231, 
		@active_start_time=220000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_客戶資料傳送]    指令碼日期: 07/30/2016 02:10:48 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:48 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_客戶資料傳送', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:48 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'delete YSBMCUST_ALL  where whsno=''LA''   

insert into YSBMCUST_ALL  
select  ''LA'' as whsno, a.*, b.custaddr1 as ctyaddr  , c.paymethod , isnull(d.payterm,30) payterm  
from  
( select custid,  custname_vn as custname , corporatecode as tbno , custshortname as custsname    from ysbmcust  ) a    
join ( select custid,  custaddr1  from ysbmaddr  ) b on b.custid = a.custid   
join ( select custid,  paymethod  from ysbmclmt) c on c.custid = a.custid  
left join ( select custid,	min(payterm) payterm from ysbdclmt where monthlysaleslimit > 0  group by custid ) d on d.custid =a.custid 
order by a.custid', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20081203, 
		@active_end_date=99991231, 
		@active_start_time=235000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_重新計算有花紅客戶AV]    指令碼日期: 07/30/2016 02:10:48 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:48 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_重新計算有花紅客戶AV', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:49 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'exec CalcAV_ForVN_HHCUST  

', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=32, 
		@freq_interval=8, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=16, 
		@freq_recurrence_factor=1, 
		@active_start_date=20110504, 
		@active_end_date=99991231, 
		@active_start_time=220000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_員工資料傳檔]    指令碼日期: 07/30/2016 02:10:49 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:49 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_員工資料傳檔', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:49 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'exec  SP_autoinsEmpfile  ', 
		@database_name=N'yfymis', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20091101, 
		@active_end_date=99991231, 
		@active_start_time=50000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_備份已送未開]    指令碼日期: 07/30/2016 02:10:49 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:49 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_備份已送未開', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'YFYMIS', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [01]    指令碼日期: 07/30/2016 02:10:50 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'exec Ins_ysbep0101  
/*
declare  @yymm as varchar(6) 
declare  @DATE2 VARCHAR(10)  
declare  @DATE3 VARCHAR(10)  

set @date3 = getdate()  
--set @date3 = ''2007/04/01''

set @yymm=convert(varchar(6),dateadd(month,-1,@date3),112) 
set @DATE2 = convert(varchar(10),dateadd(day,-1,@date3),111) 

delete  YSBEP0101 where  ym=@yymm 

Insert Into  YSBEP0101   ( deldate, do#, comp#, sc#, delqty, unitprice, m2, ym ) 
SELECT   CONVERT(char(10), a.deldate, 111) DelTime,   a.do#, a.comp#, a.sc#,  a.delqty*g.ratio  AS delqty,  k.unitprice  , 
( (ISNULL(a.delqty, 0) + ISNULL(a.additionalallowance, 0) ) ) *  case when right(a.sc#,2)<=''01'' then  x.unitarea  else ( g.[sqm/piece]* g.firstorder)  end  AS M2,  @yymm  
FROM   ysbmdord a  
inner join ysbmprod g ON a.comp# = g.comp#  
inner join ysbmaddr e ON g.invaddrno = e.addrid AND a.custid = e.custid  
inner join YSBMCUST b ON a.custid = b.custid  
inner join ysbmeord k ON a.sc# = k.sc# AND a.comp# = k.comp#  
left join  ysbmeinv M ON a.yinvno = m.yinvno  
--left join ( select * from view_ysbcp0101_unitarea ) x on x.parentcomp# = g.parentcomp# 
left join ( select  prodarea as unitarea , * from  yfypinfo ) x  on   x.parentcomp# = g.parentcomp#  
left join( select * from ysbmetax  )  z on  isnull(k.txtyp,''D'') = z.taxcode   
left join ( select * from ysbmexrt ) y on y.yyyymm=convert(char(6), k.keyindate , 112) and y.code = k.dm      
WHERE    a.status = ''1''   and a.delqty>0    and isnull(goodsstatus,'''')<>''cancel''   and right(a.sc#,2)<=''01'' 
and (  (  isnull(A.invYN,'''')=''Y''  and convert(char(6) , a.deldate, 112) < convert(char(6), m.yinvdate, 112)  and convert(char(6), m.yinvdate, 112)  > @DATE2    )    or     isnull(A.invYN,'''')<>''Y''  )  
and  left(a.sc#,1)<>''E''  and k.unitprice>0  and CONVERT(char(10), a.deldate, 111)>=''2005/09/01''       
and   CONVERT(char(10), a.deldate, 111) <= @DATE2    
and  (   convert(char(10), m.yinvdate, 111)  > @DATE2    or  isnull(convert(char(10), m.yinvdate, 111),'''')=''''   )
*/', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'01', 
		@enabled=1, 
		@freq_type=16, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=1, 
		@active_start_date=20061006, 
		@active_end_date=99991231, 
		@active_start_time=300, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_備份材質售價檔]    指令碼日期: 07/30/2016 02:10:50 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:50 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_備份材質售價檔', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:50 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare @sqlstr  as varchar(500) 
declare @yymm as varchar(6)
set @yymm = convert(char(6),getdate(),112) 
declare @tbname as varchar(50) 

begin 
	set  @tbname = ''ysbmbqty_''+@yymm 
	set @sqlstr = ''drop table '' + @tbname  
	print @sqlstr
-- 	execute(@Sqlstr) 
	set @sqlstr = ''select *  , ''+''''''''+@yymm+''''''''++'' as   ym into '' +@tbname + '' from ysbmbqty''  
	print @sqlstr
	execute(@Sqlstr) 
	
	
	set  @tbname = ''ysbmtqty_''+@yymm 
	set @sqlstr = ''drop table '' + @tbname  
	print @sqlstr
	execute(@Sqlstr) 
	set @sqlstr = ''select *  , ''+''''''''+@yymm+''''''''+'' as  ym  into '' +@tbname + '' from ysbmtqty''  
	print @sqlstr
	execute(@Sqlstr) 
end 
', 
		@database_name=N'yfymis', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=32, 
		@freq_interval=8, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=16, 
		@freq_recurrence_factor=1, 
		@active_start_date=20090105, 
		@active_end_date=99991231, 
		@active_start_time=180000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_備份每日庫存]    指令碼日期: 07/30/2016 02:10:50 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:50 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_備份每日庫存', 
		@enabled=1, 
		@notify_level_eventlog=0, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [strp01]    指令碼日期: 07/30/2016 02:10:51 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'strp01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare @bakdat as varchar(10) 
declare @nowdat as varchar(10) 
set @nowdat= getdate() 
set @bakdat = convert(char(10), dateadd(d,-1, @nowdat ), 111) 
print @bakdat
    
insert into [YSBTSTCK_Days] ( bakdat, location , subloc , palno, qty, sc#, custid, comp#, mdtm ) 
--select* from [YSBTSTCK_Days]   
select @bakdat , location , subloc , palno, qty, sc#, custid, comp#, getdate()  from YSBTSTCK    
where qty<>0 
', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20080903, 
		@active_end_date=99991231, 
		@active_start_time=110, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_備份每月授信餘額]    指令碼日期: 07/30/2016 02:10:51 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:51 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_備份每月授信餘額', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:51 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare @ym as varchar(6)
set @ym = convert(char(6), getdate()  , 112 ) 
exec  ASP_backCredit   @ym  ,'''', ''SYS''', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=32, 
		@freq_interval=8, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=16, 
		@freq_recurrence_factor=1, 
		@active_start_date=20090518, 
		@active_end_date=99991231, 
		@active_start_time=230000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_備份原紙庫存(每日)]    指令碼日期: 07/30/2016 02:10:51 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:51 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_備份原紙庫存(每日)', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:52 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'delete  ydbsntck_day where dat= convert(char(10),getdate(),111)  
insert into ydbsntck_day 
select convert(char(10),getdate(),111) , * , getdate() ,''SYS'' from  
ydbsntck', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20120826, 
		@active_end_date=99991231, 
		@active_start_time=230000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_備份期末在製品]    指令碼日期: 07/30/2016 02:10:52 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:52 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_備份期末在製品', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [strp01]    指令碼日期: 07/30/2016 02:10:52 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'strp01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'exec  Bakup_YDBNP0E01X ', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=16, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=1, 
		@active_start_date=20080402, 
		@active_end_date=99991231, 
		@active_start_time=10000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_備份期末庫存]    指令碼日期: 07/30/2016 02:10:52 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:53 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_備份期末庫存', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [01]    指令碼日期: 07/30/2016 02:10:53 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'delete ysbmstck where yymm=convert(varchar(6),dateadd(month,-1,getdate()),112)
insert into ysbmstck
select location,subloc,palno,qty,sc#,custid,comp#, convert(varchar(6),dateadd(month,-1,getdate()),112) from ysbtstck where qty<>0
UPDATE YZZMCODE  SET TBLCD = 1 where tblid = ''INV''
Update ysbdclmt Set FinMonthToDateSale = 0, currentmonth =  convert(varchar(6),getdate(),112) 
 ', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'01', 
		@enabled=1, 
		@freq_type=16, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=1, 
		@active_start_date=20061006, 
		@active_end_date=99991231, 
		@active_start_time=300, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_備份路徑單價]    指令碼日期: 07/30/2016 02:10:53 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:53 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_備份路徑單價', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [strp01]    指令碼日期: 07/30/2016 02:10:53 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'strp01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'insert into   ysbmdistX 
select  convert(varchar(6),getdate(),112) , routecode, papertype, district, 
transportfee2, mdtm, muser, getdate(), ''SYS''  
,transportfee, transportfee3  
from ysbmdist', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=16, 
		@freq_interval=25, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=1, 
		@active_start_date=20100226, 
		@active_end_date=99991231, 
		@active_start_time=233000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_替換其他字元]    指令碼日期: 07/30/2016 02:10:54 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:54 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_替換其他字元', 
		@enabled=0, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [set01]    指令碼日期: 07/30/2016 02:10:54 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'set01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'update  ysbmprod set model#=replace(model#,''?'', ''  '') 
update  ysbmprod set model_vn=replace(model_vn,''?'', ''  '') 

update  ysbmeord set model#=replace(model#,''?'', ''  '') 
update  ysbmeord set part_code=replace(part_code,''?'', '' '') 
update  ysbmeord set [p/o#]=replace([p/o#],''?'', ''  '') 
update  ysbmeord set Lot#=replace(Lot#,''?'', ''  '') 
', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=8, 
		@freq_subday_interval=1, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20061010, 
		@active_end_date=99991231, 
		@active_start_time=0, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_傳送送貨單到BD廠]    指令碼日期: 07/30/2016 02:10:54 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:54 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_傳送送貨單到BD廠', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step]    指令碼日期: 07/30/2016 02:10:55 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare  @Do As VarChar(9)    
set nocount on   
Declare YSBMDORLAT Scroll Cursor For
	select  do#  from View_SendDOToBD  group  by do#  --having  substring(do#,1,2)=''59'' 
Open YSBMDORLAT
FETCH FIRST From YSBMDORLAT into @do  
While @@Fetch_Status = 0
   BEGIN ------1
	if not exists ( select  *   from [172.22.170.33].yfymis.dbo.ysbmdord_LA   where do#=@do) 
	    begin 
		insert into [172.22.170.33].yfymis.dbo.ysbmdord_LA  
		( SC#,Item,DelDate,Trip,ParentSC#,DO#,CustID,Comp#,Lorry#,DelQty,CustDO#,GoodsStatus,AdditionalAllowance,Status,Yinvno,
		DelTime,po_qty,SqmPrice,Unitprice,Ordlength,OrdWidth,OrdMdOut,SL,RunLength,RunWidth,RunMdout,Runchops,quality,
		flute,mdscore1,mdscore2,mdscore3,mdscore4,mdscore5,addqty )  
		SELECT [sc#], [item], [deldate], [trip], [parentsc#], [do#], [custid], [comp#], [lorry#], [delqty], [model#], [goodsstatus], 
		[additionalallowance], [status], [yinvno], [deltime], [po_qty], [price], [unitprice], [length], [width], [mdout], [blankwidth],
		 [schsheetlength], [schwidth], [schmdout], [schchops], [quality], [flute], [schmdscore1], [schmdscore2], [schmdscore3], [schmdscore4],
		 [schmdscore5], [addqty] from View_SendDOToBD where do#=@do 
		update  ysbmdord set procstatus=''Y''  where do#=@do 
                       end 
	else 
      		update  ysbmdord set procstatus=''Y''  where do#=@do  
	FETCH NEXT From YSBMDORLAT  into @do   
   END -------1
Close YSBMDORLAT
DEALLOCATE YSBMDORLAT
set nocount off 
', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=4, 
		@freq_subday_interval=45, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20100821, 
		@active_end_date=99991231, 
		@active_start_time=0, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_傳送貨單資料至平政廠]    指令碼日期: 07/30/2016 02:10:55 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:55 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_傳送貨單資料至平政廠', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [01]    指令碼日期: 07/30/2016 02:10:55 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare  @Do As VarChar(9)    
set nocount on   
Declare YSBMDORLAT Scroll Cursor For
	select  do#  from View_SendDOToPZ  group  by do#  --having  substring(do#,1,2)=''59'' 
Open YSBMDORLAT
FETCH FIRST From YSBMDORLAT into @do  
While @@Fetch_Status = 0
   BEGIN ------1
	if not exists ( select  *   from [172.22.171.33].yfymis.dbo.ysbmdord_LA   where do#=@do) 
	    begin 
		insert into [172.22.171.33].yfymis.dbo.ysbmdord_LA  ( SC#,Item,DelDate,Trip,ParentSC#,DO#,CustID,Comp#,Lorry#,DelQty,CustDO#,GoodsStatus,AdditionalAllowance,Status,Yinvno,DelTime,po_qty,SqmPrice,Unitprice,Ordlength,OrdWidth,OrdMdOut,SL,RunLength,RunWidth,RunMdout,Runchops,quality )  
		select * from View_SendDOToPZ where do#=@do 
		update  ysbmdord set procstatus=''Y''  where do#=@do 
                       end 
	else 
      		update  ysbmdord set procstatus=''Y''  where do#=@do  
	FETCH NEXT From YSBMDORLAT  into @do   
   END -------1
Close YSBMDORLAT
DEALLOCATE YSBMDORLAT
set nocount off 
', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'02', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=4, 
		@freq_subday_interval=30, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20061006, 
		@active_end_date=99991231, 
		@active_start_time=70000, 
		@active_end_time=200000
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_新增每月日期]    指令碼日期: 07/30/2016 02:10:55 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:55 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_新增每月日期', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:56 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'declare @days  as int   
declare @days1  as int   
declare @days2  as int   
declare  @cDatestr as datetime 
declare  @cDatestr1 as datetime 
declare  @cDatestr2 as datetime 
set   @cDatestr= getdate()
set   @cDatestr1=  getdate()+30
set   @cDatestr2= getdate()+60 

print   @cDatestr 
print   @cDatestr1 
print   @cDatestr2 


set @days=DAY(@cDatestr+(32-DAY(@cDatestr))-DAY(@cDatestr+(32-DAY(@cDatestr))))   
set @days1=DAY(@cDatestr1+(32-DAY(@cDatestr1))-DAY(@cDatestr1+(32-DAY(@cDatestr1))))   
set @days2=DAY(@cDatestr2+(32-DAY(@cDatestr2))-DAY(@cDatestr2+(32-DAY(@cDatestr2))))   
print @days 
print @days1 
print @days2  

declare @i as int 
declare @i1 as int 
declare @i2 as int 
declare @Datestr   as  varchar(10)  
declare @Datestr1   as  varchar(10)  
declare @Datestr2   as  varchar(10)  
declare @NowMonth   as varchar(6) 
declare @NowMonth1   as varchar(6) 
declare @NowMonth2   as varchar(6) 

set   @NowMonth =  ltrim(rtrim(str(year(getdate()))))+ right(''00''+ltrim(rtrim(str(month(getdate())))) ,2) 
set   @NowMonth1 =  ltrim(rtrim(str(year(@cDatestr1))))+ right(''00''+ltrim(rtrim(str(month(@cDatestr1)))) ,2)    --right(''00''+convert(varchar(6),@cDatestr1,112) ,2)
set   @NowMonth2 =   ltrim(rtrim(str(year(@cDatestr2))))+ right(''00''+ltrim(rtrim(str(month(@cDatestr2)))) ,2)    --right(''00''+convert(varchar(6),@cDatestr2,112) ,2)
print  @NowMonth   
print  @NowMonth1   
print  @NowMonth2   
--if right(@NowMonth,2)>=''10''  
--begin  
set @i=1   
set @i1=1   
set @i2=1   

while   @i  <= @days  
   begin 
--	set @Datestr = ltrim(rtrim(str(left(@NowMonth,4))))+''/''+ltrim(rtrim(str(right(@NowMonth,2))))+''/''+right(''00''+ltrim(rtrim( str(@i))),2)  
	set @Datestr = ltrim(rtrim(str(left(@NowMonth,4))))+''/''+right(@NowMonth,2)+''/''+right(''00''+ltrim(rtrim( str(@i))),2)  
	print @Datestr 
	 set @i=@i+1    
	if not exists ( select * from  YFYCALENDR  where convert(char(10),dat,111)=convert(char(10),@Datestr,111) and convert(char(6), dat, 112) = @NowMonth  ) 
	
		insert into  YFYCALENDR  (dat, yymm) values (  @Datestr, @NowMonth ) 
    end  
	
while   @i1  <= @days1  
   begin 
	set @Datestr1 = ltrim(rtrim(str(left(@NowMonth1,4))))+''/''+right(@NowMonth1,2)+''/''+right(''00''+ltrim(rtrim( str(@i1))),2)  
	print @Datestr1 
	 set @i1=@i1+1    
	if not exists ( select * from  YFYCALENDR  where convert(char(10),dat,111)=convert(char(10),@Datestr1,111) and convert(char(6), dat, 112) = @NowMonth1  ) 
		insert into  YFYCALENDR  (dat, yymm) values (  @Datestr1, @NowMonth1 ) 
    end   
 
while   @i2  <= @days2  
   begin 
	set @Datestr2 = ltrim(rtrim(str(left(@NowMonth2,4))))+''/''+right(@NowMonth2,2)+''/''+right(''00''+ltrim(rtrim( str(@i2))),2)  
	print @Datestr2
	 set @i2=@i2+1    
	if not exists ( select * from  YFYCALENDR  where convert(char(10),dat,111)=convert(char(10),@Datestr2,111) and convert(char(6), dat, 112) = @NowMonth2  ) 
		insert into  YFYCALENDR  (dat, yymm) values (  @Datestr2, @NowMonth2 ) 
    end   
    
', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=16, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=1, 
		@active_start_date=20061101, 
		@active_end_date=99991231, 
		@active_start_time=100, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_新增送貨單及發票號碼到平政]    指令碼日期: 07/30/2016 02:10:56 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:56 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_新增送貨單及發票號碼到平政', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:56 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'insert into  [172.22.171.33].yfymis.dbo.doyinvno_la  ( do#, yinvno ) 
select  do#, yinvno  from  ysbmdord  where    
convert(char(6), deldate,112)  >=''200608'' 
and isnull(do#,'''')<>'''' and isnull(yinvno,'''')<>''''  and  isnull(goodsstatus,'''')<>''cancel''   group by  do#, yinvno 
having do# not in ( select  do# from   [172.22.171.33].yfymis.dbo.doyinvno_la  ) 
', 
		@database_name=N'YFYMIS', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch02', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20061019, 
		@active_end_date=99991231, 
		@active_start_time=30500, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [MIS_複瓦機資料回傳處理]    指令碼日期: 07/30/2016 02:10:56 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:56 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'MIS_複瓦機資料回傳處理', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [01]    指令碼日期: 07/30/2016 02:10:57 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'Declare @SC As Varchar(9) 
Declare @ITEM As  int 
Declare @fdt As Varchar(20) 
Declare CR2MISCursor Scroll Cursor For
    Select SC#,ITEM# , finishdatetime   From CR2MIS Where  ( isnull(flag ,'''') = ''''  or  isnull(flag,'''') =''R'' )  and minutes <>0 order by finishdatetime  , sc# 
Open CR2MISCursor
Fetch First From CR2MISCursor Into @SC,@ITEM, @fdt 
While @@Fetch_Status=0
Begin
	BEGIN TRANSACTION 
	exec SP_UpdateIntoItemSCR_NEW @SC,@ITEM , @fdt  
--	Update CR2MIS Set FLAG = ''*'' Where SC# = @SC and ITEM# = @ITEM
	Fetch Next From CR2MISCursor Into @SC,@ITEM, @fdt 
	If @@ERROR=0		
		COMMIT TRANSACTION 
	Else
		ROLLBACK TRANSACTION 
End
Close CR2MISCursor
Deallocate CR2MISCursor
', 
		@database_name=N'yfymis', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=4, 
		@freq_subday_interval=3, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20061006, 
		@active_end_date=99991231, 
		@active_start_time=0, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [VYFYNET_SENDCQ主管名單]    指令碼日期: 07/30/2016 02:10:57 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:57 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'VYFYNET_SENDCQ主管名單', 
		@enabled=1, 
		@notify_level_eventlog=0, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [step01]    指令碼日期: 07/30/2016 02:10:57 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'step01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'delete Tab_empfile 

insert into Tab_empfile  
select  *  
-- into Tab_empfile  
from [yfynet].dbo.view_empfile where country<>''VN''
', 
		@database_name=N'vyfynet', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'step01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=8, 
		@freq_subday_interval=3, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20130731, 
		@active_end_date=99991231, 
		@active_start_time=0, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO
/****** 物件:  Job [Vyfynet_主管資料傳送]    指令碼日期: 07/30/2016 02:10:57 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** 物件:  JobCategory [[Uncategorized (Local)]]]    指令碼日期: 07/30/2016 02:10:57 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'Vyfynet_主管資料傳送', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'沒有可用的描述。', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** 物件:  Step [sch01]    指令碼日期: 07/30/2016 02:10:58 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'sch01', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=1, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'delete Tab_empfile 

insert into Tab_empfile  
select  *  
-- into Tab_empfile  
from [yfynet].dbo.view_empfile where country<>''VN''
', 
		@database_name=N'vyfynet', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'sch01', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=8, 
		@freq_subday_interval=5, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20091021, 
		@active_end_date=99991231, 
		@active_start_time=73000, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:
