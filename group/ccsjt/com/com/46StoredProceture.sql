if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DispWrongDHZL_JP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[DispWrongDHZL_JP]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LPFXTJCreateTempPrn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LPFXTJCreateTempPrn]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LPFXTJFormSelectOption]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LPFXTJFormSelectOption]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LPFXTJPartOfTime]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LPFXTJPartOfTime]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LPFXTJSelectFighterTjPrn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LPFXTJSelectFighterTjPrn]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LPFXTJSelectPrn2ByTJ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LPFXTJSelectPrn2ByTJ]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LPFXTJSumFighterTjPrn2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LPFXTJSumFighterTjPrn2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LPFXTJSumSJByname]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LPFXTJSumSJByname]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LPFXTJSumTimeForMin]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LPFXTJSumTimeForMin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sptest]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Sptest]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TreeNet]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[TreeNet]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[createDistinctdhzl]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[createDistinctdhzl]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[verifyip]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[verifyip]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create procedure DispWrongDHZL_JP
 
 as
   if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[temptb]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
     drop table [dbo].[temptb]

   Create table temptb(
     zjh     char(20) ,
     jp      int,
     ct	     int
   )   
   
   insert into temptb 
      select distinct zjh,jp,ct=count(zjh) from dhzl group by zjh,jp
   
   select distinct zjh,jp,ct 
      from temptb,telefee where jp<>ct and temptb.zjh in (select distinct telenum from telefee)
   
   

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Encrypted object is not transferable, and script can not be generated. ******/

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


Create procedure LPFXTJFormSelectOption
   @SelectOption	varchar(50),/* @OptionName/@OptionBJ/@OptionValue */
   @SourceTable		varchar(20),
   @NewOptionName	varchar(50) output,/* Fieldname @OptionBJ Value */
   @NewOptionBJ		varchar(50) output,
   @NewOptionValue	varchar(50) output
as
/*   
   declare @SelectOption	varchar(50),
	   @SourceTable		varchar(20)  
   declare  @rid	int
   set @SelectOption='飞行总时间/>=/10'
   
   set @SourceTable='fightertj'

   exec @rid=FormSelectOption  @SelectOption , @SourceTable ,@SelectOption output
   print @SelectOption
   print @rid
*/

   declare @OptionName	varchar(20),
	   @OptionBJ	varchar(10),
	   @OptionValue	varchar(50),
	   @temp1	varchar(50),
	   @temp2	varchar(50),
	   @pos		int,
	   @Errflag	int	    /* 15  比较条件不全 */
				    /* 16 表中不存在该字段 */
				    /* 17  字段值的类型不符合要求 */
 				    /* -18 HH不是整数 */
				    /* -19 MM不是0-60的整数 */
				    /* -20 SS不是0-60的整数 */
				    /* 0 成功返回处理后的条件 */

   set @pos=charindex('/',@SelectOption)
   set @OptionName=left(@SelectOption,@pos-1)
   set @Errflag=0

   if ltrim(rtrim(@OptionName))=''   
      return 15

   /* 得到字段名后的字符串 */
   set @SelectOption=right(@SelectOption,len(@SelectOption)-@pos)
   set @pos=charindex('/',@SelectOption)

   set @OptionValue=right(@SelectOption,len(@SelectOption)-@pos)
   if ltrim(rtrim(@OptionValue))=''
      set @OptionValue='0'	

   /* 得到比较符号 */
   set @OptionBJ=rtrim(left(@SelectOption,@pos-1))
   set @temp1='NULL'

   select @temp1=Fieldname , @temp2=FieldType from Disptablefield 
      where Dispname=@OptionName and Sourcetable=@SourceTable
   /*print @temp
   print @OptionName
   print @SourceTable
   print @OptionValue*/

   if @temp1='NULL' 
      return 16

   set @OptionName=ltrim(rtrim(@temp1))

   /*判断值的类型*/
   set @temp2=ltrim(rtrim(@temp2))

   declare @hh		int,
	   @mm		int,
	   @ss		int
/*   exec @hh=LPFXTJPartOfTime '132','h'
   print @hh*/

   if @temp2='time'
      begin
        exec @hh=LPFXTJPartOfTime @OptionValue,'h'
        if @hh=-18
           return -18
        exec @mm=LPFXTJPartOfTime @OptionValue,'m'
        if @mm=-19
           return -19
    
        declare	@totalmin	int	/*在进行时间比较的时候，将时间换成分钟 */
	set @totalmin=@hh*60+@mm
        set @NewOptionValue=(str(@totalmin))
      end  

    if @temp2='num'
      begin
	if Isnumeric(@OptionValue)=0
           return 17
	set @NewOptionValue= @OptionValue
        
      end

    set @NewOptionName=@OptionName
    set @NewOptionBJ=@OptionBJ 

    return 0

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


Create procedure LPFXTJPartOfTime
    @TheTime     varchar(20), /* HH:MM or HH:MM:SS */
    @flag        varchar(2)   /* h , m ,s */

as
/*    declare @TheTime	varchar(20),
	    @flag	varchar(2)
    set @TheTime='20'
    set @flag='m'*/
    set @TheTime=rtrim(ltrim(@TheTime))
    if len(@TheTime)=0
       set @TheTime='00:00
'
    if len(@TheTime)<2
       set @TheTime='0'+@TheTime

    declare  @hh  int
    declare  @mm  int
    declare  @ss  int
    declare  @pos int

    declare  @shh	varchar(10),
	     @smm	varchar(10),
	     @sss	varchar(10)

    set @shh=''
    set @smm=''
    set @sss=''

    set @hh=0
    set @mm=0
    set @ss=0
    select @TheTime=replace(@TheTime,'：',':')
    
    set @pos=charindex(':',@TheTime)
    
    if @pos=0 
       set @shh=@TheTime /*left(@TheTime,2)*/
    else
       set @shh=left(@TheTime,@pos-1)
    
/*    print @TheTime
    print right(@TheTime,len(@TheTime)-2)
    print @shh
    print @pos*/
    
    if  charindex(':',@TheTime,@pos+1)=0
        if @pos<>0
           /* HH:MM */
           set @smm=substring(@TheTime,@pos+1,2) 
        else
           /* HHMM */
           set @smm=left(right(@TheTime,len(@TheTime)-2) ,2)

    else
        /* HH:MM:SS */
        begin
            set @TheTime = SubString(@TheTime , @pos+1 , len(@TheTime)-@pos) /* Get MM:SS */
            set @pos=charindex(':',@TheTime)
            set @smm=left(@TheTime,2)  /* MM 只有两位 */
            set @TheTime = SubString(@TheTime , @pos+1 , len(@TheTime)-@pos) /* Get SS */
            set @sss=left(@TheTime,2)  /* SS 只有两位 */
        end

    if  rtrim(@shh)<>''
        if IsNumeric(rtrim(@shh))=1
           set @hh=cast(rtrim(@shh) as int)
        else
           set @hh=-18
    else
        set @hh=0
  

    if  ltrim(@smm)<>''
        if IsNumeric(ltrim(@smm))=1
           set @mm=cast(ltrim(@smm) as int)
        else
           set @mm=-19
    else
        set @mm=0

    if  ltrim(@sss)<>''
        if IsNumeric(ltrim(@sss))=1
           set @ss=cast(ltrim(@sss) as int)
        else
           set @ss=-20
    else        set @ss=0
    
    if @mm>=60
       set @mm=-19
    if @ss>=60
       set @ss=-20

    return
       case @flag
          when 'h' then @hh	/* -1 表示小时数不为0和60之间的整数 */
          when 'm' then @mm	/* -2 表示小时数不为0和60之间的整数 */
          when 's' then @ss	/* -3 表示小时数不为0和60之间的整数 */
          else      21		/* flag is false */
       end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

Create procedure LPFXTJSelectFighterTjPrn
   @SDate      varchar(50),/* YYYY-MM-DD or YYYY-MM or YYYY  */
   @EDate      varchar(50),/* YYYY-MM-DD or YYYY-MM or YYYY  */
   @TongjiTJ   varchar(100),/* default is zqsj>='00:00' */
   @TongjiFS   varchar(5), /* y , m , d */
   @DispFS     varchar(5), /* mx , lj  */
   @TongjiObjectFlag  varchar(3), /* 0 表示具体飞行人员名字，1 表示部门 */
   @TongjiObject      varchar(8000) /* 存储飞行人员姓名或部门名称 */

as
   /*declare @SDate      varchar(50),
           @EDate      varchar(50),
           @TongjiTJ   varchar(30),
           @TongjiFS   varchar(5), 
           @TongjiObjectFlag  varchar(3),
           @TongjiObject      varchar(8000) 

   set @SDate='2003-1-1'   
   set @EDate='2003-2-28'
   set @TongjiTJ=''
   set @TongjiFS='D'
   set @TongjiObjectFlag='0'
   set @TongjiObject='蔡文建，董建平，杜义勇'*/

   if rtrim(@TongjiObject)=''
	return

   replace(@TongjiObject , ',' , '，')	
   set @TongjiObject=@TongjiObject + '，'   

   declare @startdate varchar(50)
   declare @enddate varchar(50)  

   declare @SDateOrg	varchar(50),
	   @EDateOrg	varchar(50) 

   set @SDateOrg=@SDate
   set @EDateOrg=@EDate

   exec LPFXTJCreateTempPrn

   if @TongjiFS='d'
      begin
        set @startdate=@SDate
        set @enddate=@EDate
      end 

   if @TongjiFS='m'
      begin
        set @SDate = @SDate  + '-01'
        /*避免出现大、小月及闰月的问题，即返回到上月的最末一天，再往下加一月*/
        set @enddate = (DateAdd(dd, -1, DateAdd(mm, 1, @SDate)))
       
        set @EDate = @EDate  + '-01'
        set @EDate = (DateAdd(dd, -1, DateAdd(mm, 1, @EDate)))
      end 

   if @TongjiFS='y'
      begin
        set @enddate = @SDate +  '-12-31'
        set @SDate = @SDate +  '-01-01'
        set @EDate = @EDate + '-12-31'
      end 

   set @startdate=@SDate

   while 1=1 
     Begin
       
       if cast(@enddate as smalldatetime) > cast(@EDate as smalldatetime)
          break

       /*挑选明细*/ 
       Insert into fighterTJForPrn 
          select distinct * from fighterTJ 
             where riqi>=@startdate  and riqi<=@enddate 
                   and (charindex(rtrim(pilot)+'，',@TongjiObject)>0)
        

       if @TongjiFS='d' break
       /*select sd='my'
       return*/

       /*汇总*/
       Execute LPFXTJSumFighterTjPrn2 @startdate , @enddate , @TongjiFS

       delete from fightertjforprn
       
       if @TongjiFS='m'
          begin
	    
            set @startdate=dateadd(mm , 1 , cast(@startdate as smalldatetime))
            set @enddate=dateadd(dd , -1 , dateadd(mm , 1 , @startdate))
          end

       if @TongjiFS='y'
          begin
            set @startdate=dateadd(yy , 1 , @startdate)
            set @enddate=dateadd(yy , 1 , @enddate)
          end
        

     End  /*  while 1=1 */

   if @TongjiFS<>'d' 
     begin
       /*将prn2表中的数据导入到prn表中*/
       Insert into fighterTJForPrn (pilot,riqi,yhsj,fxjlsj,zqsj,leftqlcs,rightqlcs)
          select pilot,riqi,yhsj,fxjlsj,zqsj,leftqlcs,rightqlcs from fighterTJForPrn2

       delete from fighterTJForPrn2
     end
   
   /*因为riqi字段是字符型，所以要转成 yyyy-mm-dd的格式*/
   If @TongjiFS='d'
      Update fighterTJForPrn set riqi=convert(char,convert(smalldatetime,riqi),120)
   
   /*汇总*/
   /*select sd=@startdate , ed=@enddate , tj=@TongjiFS*/
   Execute LPFXTJSumFighterTjPrn2 @startdate , @enddate , @TongjiFS

   /*对结果进行有条件的挑选*/
   Execute LPFXTJSelectPrn2ByTJ @TongjiTJ 

   /*if @DispFS='lj'*/
   update fightertjforprn2 set
             riqi=@SDateOrg + '至' + @EDateOrg

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

Create procedure LPFXTJSelectPrn2ByTJ
  @TongjiTJ	varchar(100)

As
   /*declare @TongjiTJ	varchar(20)
   set @TongjiTJ='飞行总时间/>=/60'*/
 
   Create table #fxtjtemp(
      pilot	varchar(50),
      fxjlsj	varchar(20),
      yhsj	varchar(20),
      zqsj	varchar(20),
      leftqlcs	int,
      rightqlcs int        

   )
   
   declare	@totalmin   int
   Insert into #fxtjtemp 
      select pilot,fxjlsj,yhsj,zqsj,leftqlcs,rightqlcs from fightertjforprn2

   declare tempCursor cursor for
      select pilot,fxjlsj,yhsj,zqsj,leftqlcs,rightqlcs from #fxtjtemp 
      for update

   declare @pilot     varchar(20),
	   @fxjlsj    varchar(20),
           @yhsj      varchar(20),
           @zqsj      varchar(20),
           @leftqlcs	int,
	   @rightqlcs	int
   open tempCursor
   fetch  next from tempCursor into @pilot,@fxjlsj,@yhsj,@zqsj,@leftqlcs,@rightqlcs

   while ( @@fetch_status = 0) 
     begin
       /* 将时间HH:MM转化成分钟值 */
       Execute  LPFXTJSumTimeForMin -1,'00:00',@fxjlsj,'2',@fxjlsj output
       Execute  LPFXTJSumTimeForMin -1,'00:00',@yhsj,'2',@yhsj output
       Execute  LPFXTJSumTimeForMin -1,'00:00',@zqsj,'2',@zqsj output

       Update #fxtjtemp set
          fxjlsj=@fxjlsj,zqsj=@zqsj,yhsj=@yhsj
          where current of tempCursor

       Fetch  next from tempCursor into @pilot,@fxjlsj,@yhsj,@zqsj,@leftqlcs,@rightqlcs       
     end

   close tempCursor
   deallocate tempCursor
    
   declare   @NewOptionName		varchar(30),
	     @NewOptionBJ		varchar(30),
	     @NewOptionValue		varchar(30)

   /* 形成筛选结果的条件 */
   exec LPFXTJFormSelectOption @TongjiTJ ,'fightertj', @NewOptionName output, @NewOptionBJ output, @NewOptionValue output
   
   set @NewOptionName=rtrim(ltrim(@NewOptionName))   
   set @NewOptionBJ=rtrim(ltrim(@NewOptionBJ))
   set @NewOptionValue=rtrim(ltrim(@NewOptionValue))

   /*print   @NewOptionName
   print   @NewOptionBJ		
   print   @NewOptionValue*/

   if @NewOptionBJ='>'
     delete from fightertjforprn2
        where pilot not In
         ( 
	   /*declare  @NewOptionValue	varchar(20),
		    @NewOptionName	varchar(20)
	   set @NewOptionValue='60'
	   set @NewOptionName='zqsj'

           select * from #fxtjtemp where cast(zqsj as int)> cast(@NewOptionValue as int)*/

           select pilot from #fxtjtemp where 
             cast (
             case @NewOptionName 
                  when 'fxjlsj' then fxjlsj
                  when 'zqsj'   then zqsj
		  when 'yhsj'   then yhsj
		  when 'leftqlcs' then leftqlcs
                  when 'rightqlcs' then rightqlcs
             end 
             as int) > cast(@NewOptionValue as int)
         )  

   if @NewOptionBJ='<'
     delete from fightertjforprn2
        where pilot not In
         ( select pilot from #fxtjtemp where 
	     cast (
             case @NewOptionName 
                  when 'fxjlsj' then fxjlsj
                  when 'zqsj'   then zqsj
		  when 'yhsj'   then yhsj
		  when 'leftqlcs' then leftqlcs
                  when 'rightqlcs' then rightqlcs
             end 
             as int) < cast(@NewOptionValue as int)
         )  

   if @NewOptionBJ='='
     delete from fightertjforprn2
        where pilot not In
         ( 

           select pilot from #fxtjtemp where 
             cast (
             case @NewOptionName 
                  when 'fxjlsj' then fxjlsj
                  when 'zqsj'   then zqsj
		  when 'yhsj'   then yhsj
		  when 'leftqlcs' then leftqlcs
                  when 'rightqlcs' then rightqlcs
             end 
             as int) = cast(@NewOptionValue as int)
         ) 
 
   if @NewOptionBJ='>='
     delete from fightertjforprn2
        where pilot not In
         ( 

           select pilot from #fxtjtemp where 
             cast (
             case @NewOptionName 
                  when 'fxjlsj' then fxjlsj
                  when 'zqsj'   then zqsj
		  when 'yhsj'   then yhsj
		  when 'leftqlcs' then leftqlcs
                  when 'rightqlcs' then rightqlcs
             end 
             as int) >= cast(@NewOptionValue as int)
         )  

   if @NewOptionBJ='<='
     delete from fightertjforprn2
        where pilot not In
         ( 

           select pilot from #fxtjtemp where 
             cast (
             case @NewOptionName 
                  when 'fxjlsj' then fxjlsj
                  when 'zqsj'   then zqsj
		  when 'yhsj'   then yhsj
		  when 'leftqlcs' then leftqlcs
                  when 'rightqlcs' then rightqlcs
             end 
             as int) <= cast(@NewOptionValue as int)
         )  

   delete from fightertjforprn
         where pilot not in (select distinct pilot from fightertjforprn2 )

   drop table #fxtjtemp
   

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

Create procedure LPFXTJSumFighterTjPrn2 /*根据@TongjiFS对 Prn表中所有数据进行统计 */
   @startdate        varchar(50),
   @enddate          varchar(50), /*在最后一次的统计中，该日期无用处*/
   @TongjiFS         varchar(3)
   
as
   /*declare @startdate        varchar(50),
           @enddate          varchar(50),
           @TongjiFS         varchar(3)
   set @startdate='2003-1-1'
   set @enddate= '2003-2-28'
   set @TongjiFS='d'
      
   set @startdate=convert(char,convert(smalldatetime,@startdate),120)
   set @enddate=convert(char,convert(smalldatetime,@enddate),120)
   print @startdate
   print str(Month(@enddate))*/
   
   declare @pilot varchar(50),
           @riqi  varchar(50)
   declare @fetchstatusTemp   int
   
   if @TongjiFS = 'd'
     declare PrnCursor scroll cursor for 
             select distinct pilot,riqi='' from fighterTJForPrn 
   if @TongjiFS = 'm'
     declare PrnCursor scroll cursor for 
             select distinct pilot,riqi= ltrim(rtrim(str(Year(@enddate)))) + '-' + ltrim(rtrim(str(Month(@enddate)))) from fighterTJForPrn 
   if @TongjiFS = 'y'
     declare PrnCursor scroll cursor for 
             select distinct pilot,riqi= ltrim(rtrim(str(Year(@enddate)))) from fighterTJForPrn 

   open PrnCursor
   fetch next from PrnCursor into @pilot,@riqi /* It will locate the first record */
   
  declare @i  int
  set @i=1
   while ( @@fetch_status = 0) 
     begin
        /*print @i
        print @pilot
        print @riqi
        print @startdate
        print @enddate*/
        exec  LPFXTJSumSJByname @pilot,@riqi,@startdate,@enddate
        
        fetch next from PrnCursor into @pilot,@riqi
        set @i=@i+1
     end

   close PrnCursor
   deallocate PrnCursor
   

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

Create procedure LPFXTJSumSJByname
   @pilot		varchar(20),
   @riqi		varchar(50),
   @startdate		varchar(50),
   @enddate		varchar(50)
   /*@Fieldname		varchar(200),
   @SourceTable		varchar(30),
   @DestTable		varchar(30),*//* 如果 len()=0 then result is in @ReturnValue */
/*   @ReturnValue		varchar(200)	output*/
As
/*   declare
     @pname		varchar(20),
     @Fieldname		varchar(200),
     @SourceTable		varchar(30),
     @DestTable		varchar(30), 
     @ReturnValue		varchar(200) 
   
   set @pname='杜义勇'  
   set @Fieldname='pilot,fxjlsj,zqsj,yhsj,leftqlcs,rightqlcs'
   set @SourceTable='fightertjforprn'
   set @destTable='fightertjforprn2'
*/
/*   declare @pilot	varchar(20) ,
	   @riqi	varchar(20) ,
	   @startdate	varchar(20) ,
	   @enddate	varchar(20) 
   set @pilot = '蔡文建'
   set @riqi  = ''
   set @startdate = '2003-1-1'
   set @enddate	 = '2003-2-28'
*/
   declare @sumfxjlsj varchar(20),
           @sumyhsj   varchar(20),
           @sumzqsj   varchar(20),
           @sumleftqlcs       int,
           @sumrightqlcs      int,
           @fxjlsj    varchar(20),
           @yhsj      varchar(20),
           @zqsj      varchar(20),
           @leftqlcs          int,
           @rightqlcs         int

   set @sumfxjlsj='00:00'
   set @sumyhsj='00:00'
   set @sumzqsj='00:00'
   set @sumleftqlcs=0
   set @sumrightqlcs=0 

   declare SelCursor scroll cursor for
           select fxjlsj=IsNull(fxjlsj,'00:00'),yhsj=IsNull(yhsj,'00:00'),zqsj=IsNull(zqsj,'00:00'),leftqlcs=IsNull(leftqlcs,0),rightqlcs=IsNull(rightqlcs,0) 
              from fighterTJForPrn where pilot = @pilot

   open SelCursor
   fetch next from SelCursor into @fxjlsj , @yhsj , @zqsj , @leftqlcs , @rightqlcs
   declare @i  int
   set @i=0
   while ( @@fetch_status = 0)       
           begin
              set @sumfxjlsj=ltrim(rtrim(@sumfxjlsj))
              set @sumyhsj=ltrim(rtrim(@sumyhsj))
	      set @sumzqsj=ltrim(rtrim(@sumzqsj))
              set @fxjlsj=ltrim(rtrim(@fxjlsj))
              set @yhsj=ltrim(rtrim(@yhsj))
	      set @zqsj=ltrim(rtrim(@zqsj))

              /*if @i=3
                 begin
                   select jlsj=@sumfxjlsj , yhsj=@sumyhsj , zqsj=@sumzqsj ,lef=@sumleftqlcs
                   return
                 end
              
	      print @i
              print @sumfxjlsj
              print @fxjlsj
              print @sumyhsj
              print @yhsj
              print @sumzqsj
              print @zqsj  */

              execute LPFXTJSumTimeForMin -1,@sumfxjlsj,@fxjlsj,'3', @sumfxjlsj output
              execute LPFXTJSumTimeForMin -1,@sumyhsj,@yhsj,'3', @sumyhsj output
              execute LPFXTJSumTimeForMin -1,@sumzqsj,@zqsj,'3', @sumzqsj output

              set @sumleftqlcs = @sumleftqlcs + @leftqlcs
              set @sumrightqlcs = @sumrightqlcs + @rightqlcs
              
              fetch next from SelCursor into @fxjlsj , @yhsj , @zqsj , @leftqlcs , @rightqlcs
              set @i=@i+1
           end 
        
   Insert into fightertjforprn2 
           (pilot,riqi,startdate,enddate,yhsj,fxjlsj,zqsj,leftqlcs,rightqlcs) 
           values( rtrim(@pilot) , rtrim(@riqi) , @startdate , @enddate , rtrim(@sumyhsj) , rtrim(@sumfxjlsj) , rtrim(@sumzqsj) , @sumleftqlcs , @sumrightqlcs )
        
   close SelCursor
   deallocate SelCursor


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

Create procedure LPFXTJSumTimeForMin
    @TotalMin      int,        /* Min value , if =-1 then select @TotalTime */
    @TotalTime     varchar(20),/* HH:MM */
    @SingleTime    varchar(20),/* HH:MM */
    @ReturnFlag    varchar(10),/* 1表示返回分钟值以RETURN方式(整数型)，
                                  2表示返回分钟值以 @ReturnValue 方式，
                                  3表示返回HH:MM格式以 @ReturnValue 方式，
                                  4表示返回带小数的小时以 @ReturnValue 方式*/
    @ReturnValue   varchar(20)= '00:00' output 
AS
    declare @flag int
    declare  @hh  int
    declare  @mm  int
    declare  @mmstr  varchar(10)
    
    set @TotalTime=ltrim(rtrim(@TotalTime))
    set @SingleTime=ltrim(rtrim(@SingleTime))

    execute @hh=LPFXTJPartOfTime @SingleTime , 'h'
    execute @mm=LPFXTJPartOfTime @SingleTime , 'm'

    set @flag = @TotalMin
    if @flag = -1   /* 此时应该从@TotalTime取总值 */ 
       begin
          declare @thh  int
          declare @tmm  int
          exec @thh = LPFXTJPartOfTime @TotalTime , 'h' 
          exec @tmm = LPFXTJPartOfTime @TotalTime , 'm'        
          set @TotalMin = @thh * 60 + @tmm 
       end

    select @TotalMin = @TotalMin + @hh*60 + @mm
    
    if IsNumeric(@TotalMin)=0
       set @TotalMin=0

    if @ReturnFlag='1'
       return @TotalMin
    if @ReturnFlag='2' 
       set @ReturnValue = str(@TotalMin)
    if @ReturnFlag='3'
       begin
        set @mmstr = cast(@TotalMin%60 as varchar(10))
        if len(@mmstr)<2  /*8:5 is 8:05 not 8:50*/
           set @ReturnValue = cast(@TotalMin/60 as varchar(12)) + ':0' + @mmstr
        else
           set @ReturnValue = cast(@TotalMin/60 as varchar(12)) + ':' + @mmstr
       end  
    if @ReturnFlag='4' 
       set @ReturnValue = str(@TotalMin/60.00)
    return


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE Sptest AS

select * from test
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE TreeNet AS
declare @cnt smallint
Select AgentID into #leaves from AgentTree where AgentID not in (select distinct FAgentID from AgentTree)
Select distinct FAgentID into #roots from AgentTree where FAgentID not in (select AgentID from AgentTree)
select AgentID,FAgentID into #source from AgentTree
  /* where AgentID in (select AgentID from #leaves) */
select AgentID,AgentID As FAgentID into #dest from AgentTree 
insert #dest select FAgentID,FAgentID from #roots 
select @cnt=(select count(*) from #leaves)
while @cnt>0 
begin
   insert #dest select AgentID,FAgentID from #source
   delete #source where FAgentID in (select FAgentID from #roots)
   select @cnt=(select count(*) from #source)
   update #source set FAgentID=b.FAgentID from #source,AgentTree b where #source.FAgentID=b.AgentID
end
delete from AgentNet
insert AgentNet (AgentID,ManageAgentID) select * from #dest order by AgentID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



create procedure createDistinctdhzl
     @flag   int
as
  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[distinct_dhzl]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
     drop table [dbo].[distinct_dhzl]
  
  Create table distinct_dhzl(
     CountID int not null identity,
     zjh     char(20) ,
     yhmc    char(100),
     bz      char(100),
     jp      int
  )

  if @flag=0
  begin
     /*是按代码排序，因此一个用户名称对应一个电话*/
     Insert into distinct_dhzl select distinct zjh,yhmc,bz,jp from dhzl
     return
  end
  
  /*没有按代码排序，一个电话对应一系列的用户名称，中间用逗号相隔*/
  /*默认的@flag值为1 */
  Insert into distinct_dhzl select distinct zjh,'','',1 from dhzl
    
  declare dhzlCursor scroll cursor 
     for  select distinct zjh,yhmc,bz from dhzl

  declare @zjh   char(20),
          @yhmc  char(100),
          @bz    char(100)
  
  set @yhmc=''
  set @bz='' 
  open dhzlCursor
  fetch next from dhzlCursor into @zjh,@yhmc,@bz
  
  while (@@fetch_status = 0) 
  begin
    
    /*因为是select 语句，所以返回的结果一般都是带有右空格，所以下面的三句的rtrim并没有用*/
    select @zjh=ltrim(rtrim(IsNull(@zjh,' ')))
    select @yhmc=ltrim(rtrim(IsNull(@yhmc,' ')))
    select @bz=ltrim(rtrim(IsNull(@bz,' ')))

    if @zjh<>''
       update distinct_dhzl set
           yhmc=rtrim(yhmc) + rtrim(@yhmc) +','  , bz=rtrim(bz)  + rtrim(@bz) +','
           where rtrim(zjh)=rtrim(@zjh)
    fetch next from dhzlCursor into @zjh,@yhmc,@bz

  end /*  end of dhzlCursor cursor loop */

  /*清除末尾的逗号*/
  update distinct_dhzl set 
           yhmc=left(rtrim(yhmc),len(rtrim(yhmc))-1)
           where right(rtrim(yhmc),1)=','

  update distinct_dhzl set 
           bz=left(rtrim(bz),len(rtrim(bz))-1)
           where right(rtrim(bz),1)=','
  close dhzlCursor
  deallocate dhzlCursor

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


create procedure verifyip
    @remoteaddr    varchar(20)
as
    declare  @temp varchar(20)
    declare  @tp varchar(20)
    declare  @i  int
    declare  @j  int

    select @temp=@remoteaddr+'.'
    select @i=0

    while @i<=3

    begin
   	select @temp=left(@temp,len(@temp)-charindex('.',reverse(@temp)))
	select @tp=@temp
   	select @j=@i
   	while @j>0  
     	  begin
            select @tp=@tp+'.*'
            select @j=@j-1
          end
       select @j=count(*) from AllowIPTable where AllowIP=@tp      

       if @j=1	
         begin 
          select @j 
          return
         end

       select @i=@i+1
       if @i=0 
         break
       else
         continue
    end
    select 0
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

