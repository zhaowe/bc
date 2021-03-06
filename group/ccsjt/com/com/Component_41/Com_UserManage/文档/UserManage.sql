CREATE TABLE [dbo].[CreditLog] (
	[LogID] [uniqueidentifier] NOT NULL ,
	[CreditID] [char] (10) NULL ,
	[AgentID] [uniqueidentifier] NULL ,
	[OperationType] [char] (2) NULL ,
	[CreditChange] [money] NULL ,
	[OrderID] [uniqueidentifier] NULL ,
	[LogNote] [nvarchar] (100) NULL ,
	[Operator] [varchar] (50) NULL ,
	[OperationTime] [datetime] NOT NULL ,
	[Memo] [varchar] (100) NULL ,
	[SettleStatus] [char] (1) NOT NULL ,
	[SettleTime] [datetime] NULL ,
	[SettleID] [uniqueidentifier] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AgentInfo] (
	[AgentID] [uniqueidentifier] NOT NULL ,
	[AgentOffice] [char] (10) NOT NULL ,
	[AgentType] [char] (1) NULL ,
	[AgentName] [nvarchar] (100) NULL ,
	[AgentShortName] [nvarchar] (100) NULL ,
	[UseObject] [char] (5) NULL ,
	[AgentCity] [nvarchar] (100) NULL ,
	[lxrAdd] [nvarchar] (100) NULL ,
	[lxrPhone] [nvarchar] (100) NULL ,
	[lxrName] [nvarchar] (100) NULL ,
	[OpenBank] [nvarchar] (100) NULL ,
	[OpenAccount] [nvarchar] (50) NULL ,
	[ProtocalNo] [nvarchar] (20) NULL ,
	[ProtocalDate] [datetime] NULL ,
	[Operator] [varchar] (50) NULL ,
	[CreateTime] [datetime] NOT NULL ,
	[EndTime] [datetime] NULL ,
	[AgentEntity] [char] (1) NULL ,
	[DealStatus] [char] (1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AgentJZ] (
	[ID] [uniqueidentifier] NOT NULL ,
	[BookAgentID] [uniqueidentifier] NOT NULL ,
	[TicketOutAgentID] [uniqueidentifier] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AgentNet] (
	[ID] [uniqueidentifier] NOT NULL ,
	[AgentID] [uniqueidentifier] NOT NULL ,
	[ManageAgentID] [uniqueidentifier] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AgentToCredit] (
	[FlowID] [uniqueidentifier] NOT NULL ,
	[AgentID] [uniqueidentifier] NOT NULL ,
	[CreditID] [char] (10) NOT NULL ,
	[ReceiptTime] [datetime] NOT NULL ,
	[Operator] [varchar] (50) NULL ,
	[EndTime] [datetime] NULL ,
	[ValidFlag] [char] (1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AgentTree] (
	[ID] [uniqueidentifier] NOT NULL ,
	[AgentID] [uniqueidentifier] NOT NULL ,
	[FAgentID] [uniqueidentifier] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CompanyInfo] (
	[CompanyId] [char] (5) NOT NULL ,
	[Description] [nvarchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CompanyLocale] (
	[CompanyID] [char] (5) NOT NULL ,
	[Locale] [char] (5) NOT NULL ,
	[CompanyName] [nvarchar] (50) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CreditAccount] (
	[CreditID] [char] (10) NOT NULL ,
	[DepositCash] [money] NOT NULL ,
	[CreditCash] [money] NOT NULL ,
	[TotalCreditAmt] [money] NOT NULL ,
	[CreditPriority] [char] (1) NULL ,
	[CashType] [char] (3) NULL ,
	[CompanyName] [nvarchar] (100) NULL ,
	[UseObject] [char] (5) NULL ,
	[CreditCity] [nvarchar] (100) NULL ,
	[lxrphone] [nvarchar] (100) NULL ,
	[lxrName] [nvarchar] (100) NULL ,
	[lxrAdd] [nvarchar] (100) NULL ,
	[OpenBank] [nvarchar] (100) NULL ,
	[OpenAccount] [nvarchar] (20) NULL ,
	[ProtocalNo] [nvarchar] (100) NULL ,
	[ProtocalDate] [datetime] NULL ,
	[ReceiptTime] [datetime] NOT NULL ,
	[Operator] [varchar] (50) NULL ,
	[EndTime] [datetime] NULL ,
	[ValidFlag] [char] (1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CreditSettleLog] (
	[SettleID] [uniqueidentifier] NOT NULL ,
	[CreditID] [char] (10) NULL ,
	[StartDate] [datetime] NULL ,
	[EndDate] [datetime] NULL ,
	[TotalRetCredit] [money] NULL ,
	[PriceRequired] [money] NULL ,
	[SupplyCash] [money] NULL ,
	[Operator] [varchar] (50) NULL ,
	[SettleTime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Function] (
	[FunctionID] [uniqueidentifier] NOT NULL ,
	[Description] [varchar] (50) NOT NULL ,
	[OrderNum] [int] NOT NULL ,
	[Conflict] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FunctionConflict] (
	[Conflict1] [int] NOT NULL ,
	[Conflict2] [int] NOT NULL ,
	[Description] [nvarchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FunctionLocale] (
	[FunctionID] [uniqueidentifier] NOT NULL ,
	[Locale] [char] (5) NOT NULL ,
	[FunctionName] [nvarchar] (50) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[GroupFunction] (
	[GroupID] [uniqueidentifier] NOT NULL ,
	[FunctionID] [uniqueidentifier] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[GroupInfo] (
	[GroupID] [uniqueidentifier] NOT NULL ,
	[Description] [nvarchar] (50) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[GroupLocale] (
	[GroupID] [uniqueidentifier] NOT NULL ,
	[Locale] [char] (5) NOT NULL ,
	[GroupName] [nvarchar] (50) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Locale] (
	[Locale] [char] (5) NOT NULL ,
	[LocaleName] [varchar] (50) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LoginInfo] (
	[LoginID] [nvarchar] (20) NOT NULL ,
	[Name] [nvarchar] (50) NULL ,
	[Sex] [char] (1) NULL ,
	[AgentID] [uniqueidentifier] NOT NULL ,
	[CompanyID] [char] (5) NOT NULL ,
	[ContactInfo] [nvarchar] (50) NULL ,
	[UseObject] [char] (5) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UseObject] (
	[UseObject] [char] (5) NOT NULL ,
	[UseObjectName] [varchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserFunction] (
	[UserID] [uniqueidentifier] NOT NULL ,
	[FunctionID] [uniqueidentifier] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserGroup] (
	[UserID] [uniqueidentifier] NOT NULL ,
	[GroupID] [uniqueidentifier] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserHistory] (
	[UserID] [uniqueidentifier] NOT NULL ,
	[LoginID] [nvarchar] (20) NOT NULL ,
	[StartDate] [datetime] NOT NULL ,
	[EndDate] [datetime] NULL ,
	[Status] [char] (1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserInfo] (
	[UserID] [uniqueidentifier] NOT NULL ,
	[LoginID] [nvarchar] (20) NOT NULL ,
	[Password] [varchar] (10) NOT NULL ,
	[EndDate] [datetime] NULL ,
	[Status] [char] (1) NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[CreditLog] WITH NOCHECK ADD 
	CONSTRAINT [DF_CreditLog_OperationTime] DEFAULT (getdate()) FOR [OperationTime],
	CONSTRAINT [DF_CreditLog_SettleStatus] DEFAULT ('N') FOR [SettleStatus],
	CONSTRAINT [PK_CreditLog] PRIMARY KEY  NONCLUSTERED 
	(
		[LogID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AgentInfo] WITH NOCHECK ADD 
	CONSTRAINT [DF_AgentInfo_AgentID] DEFAULT (newid()) FOR [AgentID],
	CONSTRAINT [DF_AgentInfo_UseObject] DEFAULT ('AMS') FOR [UseObject],
	CONSTRAINT [DF_AgentInfo_CreateTime] DEFAULT (getdate()) FOR [CreateTime],
	CONSTRAINT [DF_SaleAgentInfo_NormalFlag] DEFAULT ('V') FOR [DealStatus],
	CONSTRAINT [PK_AgentInfo] PRIMARY KEY  NONCLUSTERED 
	(
		[AgentID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AgentJZ] WITH NOCHECK ADD 
	CONSTRAINT [DF_AgentJZ_ID] DEFAULT (newid()) FOR [ID],
	CONSTRAINT [PK_AgentJZ] PRIMARY KEY  NONCLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AgentNet] WITH NOCHECK ADD 
	CONSTRAINT [DF_AgentNet_ID] DEFAULT (newid()) FOR [ID],
	CONSTRAINT [PK_AgentNet] PRIMARY KEY  NONCLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AgentToCredit] WITH NOCHECK ADD 
	CONSTRAINT [DF_AgentToCredit_FlowID] DEFAULT (newid()) FOR [FlowID],
	CONSTRAINT [DF_AgentToCredit_ReceiptTime] DEFAULT (getdate()) FOR [ReceiptTime],
	CONSTRAINT [DF_AgentToCredit_ValidFlag] DEFAULT ('Y') FOR [ValidFlag],
	CONSTRAINT [PK_AgentToCredit] PRIMARY KEY  NONCLUSTERED 
	(
		[FlowID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AgentTree] WITH NOCHECK ADD 
	CONSTRAINT [DF_AgentTree_ID] DEFAULT (newid()) FOR [ID],
	CONSTRAINT [PK_AgentTree] PRIMARY KEY  NONCLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CompanyInfo] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyInfo] PRIMARY KEY  NONCLUSTERED 
	(
		[CompanyId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CompanyLocale] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyLocale] PRIMARY KEY  NONCLUSTERED 
	(
		[CompanyID],
		[Locale]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CreditAccount] WITH NOCHECK ADD 
	CONSTRAINT [DF_CreditAccount_DepositCash] DEFAULT (0) FOR [DepositCash],
	CONSTRAINT [DF_CreditAccount_CreditCash] DEFAULT (0) FOR [CreditCash],
	CONSTRAINT [DF_CreditAccount_TotalCreditAmt] DEFAULT (0) FOR [TotalCreditAmt],
	CONSTRAINT [DF_CreditAccount_CreditPriority] DEFAULT ('A') FOR [CreditPriority],
	CONSTRAINT [DF_CreditAccount_CashType] DEFAULT ('RMB') FOR [CashType],
	CONSTRAINT [DF_CreditAccount_ReceiptTime] DEFAULT (getdate()) FOR [ReceiptTime],
	CONSTRAINT [DF_CreditAccount_ValidFlag] DEFAULT ('Y') FOR [ValidFlag],
	CONSTRAINT [PK_CreditAccount] PRIMARY KEY  NONCLUSTERED 
	(
		[CreditID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CreditSettleLog] WITH NOCHECK ADD 
	CONSTRAINT [PK_CreditLogSettleDetail] PRIMARY KEY  NONCLUSTERED 
	(
		[SettleID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Function] WITH NOCHECK ADD 
	CONSTRAINT [DF_Function_FunctionID] DEFAULT (newid()) FOR [FunctionID],
	CONSTRAINT [PK_Function] PRIMARY KEY  NONCLUSTERED 
	(
		[FunctionID]
	)  ON [PRIMARY] ,
	CONSTRAINT [IX_Function] UNIQUE  NONCLUSTERED 
	(
		[OrderNum]
	)  ON [PRIMARY] ,
	CONSTRAINT [IX_Function_1] UNIQUE  NONCLUSTERED 
	(
		[Description]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FunctionConflict] WITH NOCHECK ADD 
	CONSTRAINT [PK_FunctionConflict] PRIMARY KEY  NONCLUSTERED 
	(
		[Conflict1],
		[Conflict2]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FunctionLocale] WITH NOCHECK ADD 
	CONSTRAINT [PK_FunctionLocale] PRIMARY KEY  NONCLUSTERED 
	(
		[FunctionID],
		[Locale]
	)  ON [PRIMARY] ,
	CONSTRAINT [IX_FunctionLocale] UNIQUE  NONCLUSTERED 
	(
		[FunctionName]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[GroupFunction] WITH NOCHECK ADD 
	CONSTRAINT [PK_GroupFunction] PRIMARY KEY  NONCLUSTERED 
	(
		[GroupID],
		[FunctionID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[GroupInfo] WITH NOCHECK ADD 
	CONSTRAINT [DF_GroupInfo_GroupID] DEFAULT (newid()) FOR [GroupID],
	CONSTRAINT [PK_GroupInfo] PRIMARY KEY  NONCLUSTERED 
	(
		[GroupID]
	)  ON [PRIMARY] ,
	CONSTRAINT [IX_GroupInfo] UNIQUE  NONCLUSTERED 
	(
		[Description]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[GroupLocale] WITH NOCHECK ADD 
	CONSTRAINT [PK_GroupLocale] PRIMARY KEY  NONCLUSTERED 
	(
		[GroupID],
		[Locale]
	)  ON [PRIMARY] ,
	CONSTRAINT [IX_GroupLocale] UNIQUE  NONCLUSTERED 
	(
		[GroupName]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Locale] WITH NOCHECK ADD 
	CONSTRAINT [PK_Locale] PRIMARY KEY  NONCLUSTERED 
	(
		[Locale]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[LoginInfo] WITH NOCHECK ADD 
	CONSTRAINT [PK_LoginInfo] PRIMARY KEY  NONCLUSTERED 
	(
		[LoginID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UseObject] WITH NOCHECK ADD 
	CONSTRAINT [PK_UseObject] PRIMARY KEY  NONCLUSTERED 
	(
		[UseObject]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserFunction] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserFunction] PRIMARY KEY  NONCLUSTERED 
	(
		[UserID],
		[FunctionID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserGroup] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserGroup] PRIMARY KEY  NONCLUSTERED 
	(
		[UserID],
		[GroupID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserHistory] WITH NOCHECK ADD 
	CONSTRAINT [DF_UserHistory_Status] DEFAULT ('E') FOR [Status],
	CONSTRAINT [PK_UserHistory] PRIMARY KEY  NONCLUSTERED 
	(
		[UserID],
		[StartDate]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserInfo] WITH NOCHECK ADD 
	CONSTRAINT [DF_UserInfo_UserID] DEFAULT (newid()) FOR [UserID],
	CONSTRAINT [DF_UserInfo_Status] DEFAULT ('E') FOR [Status],
	CONSTRAINT [PK_UserInfo] PRIMARY KEY  NONCLUSTERED 
	(
		[UserID]
	)  ON [PRIMARY] ,
	CONSTRAINT [IX_UserInfo] UNIQUE  NONCLUSTERED 
	(
		[LoginID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CompanyLocale] ADD 
	CONSTRAINT [FK_CompanyLocale_CompanyInfo] FOREIGN KEY 
	(
		[CompanyID]
	) REFERENCES [dbo].[CompanyInfo] (
		[CompanyId]
	)
GO

ALTER TABLE [dbo].[FunctionLocale] ADD 
	CONSTRAINT [FK_FunctionLocale_Function] FOREIGN KEY 
	(
		[FunctionID]
	) REFERENCES [dbo].[Function] (
		[FunctionID]
	),
	CONSTRAINT [FK_FunctionLocale_Locale] FOREIGN KEY 
	(
		[Locale]
	) REFERENCES [dbo].[Locale] (
		[Locale]
	)
GO

ALTER TABLE [dbo].[GroupFunction] ADD 
	CONSTRAINT [FK_GroupFunction_Function] FOREIGN KEY 
	(
		[FunctionID]
	) REFERENCES [dbo].[Function] (
		[FunctionID]
	),
	CONSTRAINT [FK_GroupFunction_GroupInfo] FOREIGN KEY 
	(
		[GroupID]
	) REFERENCES [dbo].[GroupInfo] (
		[GroupID]
	)
GO

ALTER TABLE [dbo].[GroupLocale] ADD 
	CONSTRAINT [FK_GroupLocale_GroupInfo] FOREIGN KEY 
	(
		[GroupID]
	) REFERENCES [dbo].[GroupInfo] (
		[GroupID]
	),
	CONSTRAINT [FK_GroupLocale_Locale] FOREIGN KEY 
	(
		[Locale]
	) REFERENCES [dbo].[Locale] (
		[Locale]
	)
GO

ALTER TABLE [dbo].[LoginInfo] ADD 
	CONSTRAINT [FK_LoginInfo_AgentInfo] FOREIGN KEY 
	(
		[AgentID]
	) REFERENCES [dbo].[AgentInfo] (
		[AgentID]
	),
	CONSTRAINT [FK_LoginInfo_CompanyInfo] FOREIGN KEY 
	(
		[CompanyID]
	) REFERENCES [dbo].[CompanyInfo] (
		[CompanyId]
	),
	CONSTRAINT [FK_LoginInfo_UseObject] FOREIGN KEY 
	(
		[UseObject]
	) REFERENCES [dbo].[UseObject] (
		[UseObject]
	)
GO

ALTER TABLE [dbo].[UserFunction] ADD 
	CONSTRAINT [FK_UserFunction_Function] FOREIGN KEY 
	(
		[FunctionID]
	) REFERENCES [dbo].[Function] (
		[FunctionID]
	),
	CONSTRAINT [FK_UserFunction_UserInfo] FOREIGN KEY 
	(
		[UserID]
	) REFERENCES [dbo].[UserInfo] (
		[UserID]
	)
GO

ALTER TABLE [dbo].[UserGroup] ADD 
	CONSTRAINT [FK_UserGroup_GroupInfo] FOREIGN KEY 
	(
		[GroupID]
	) REFERENCES [dbo].[GroupInfo] (
		[GroupID]
	),
	CONSTRAINT [FK_UserGroup_UserInfo] FOREIGN KEY 
	(
		[UserID]
	) REFERENCES [dbo].[UserInfo] (
		[UserID]
	)
GO

ALTER TABLE [dbo].[UserInfo] ADD 
	CONSTRAINT [FK_UserInfo_LoginInfo] FOREIGN KEY 
	(
		[LoginID]
	) REFERENCES [dbo].[LoginInfo] (
		[LoginID]
	)
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.rpt_AgentCreditInfo
AS
SELECT AgentInfo.AgentOffice, AgentToCredit.CreditID, AgentInfo.AgentType, 
      AgentInfo.AgentName, AgentInfo.AgentShortName, AgentInfo.lxrAdd, 
      AgentInfo.lxrPhone, AgentInfo.lxrName, AgentToCredit.ValidFlag, 
      CreditAccount.DepositCash, CreditAccount.CreditCash, CreditAccount.TotalCreditAmt, 
      AgentToCredit.AgentID, AgentInfo.DealStatus
FROM AgentInfo INNER JOIN
      AgentToCredit ON AgentInfo.AgentID = AgentToCredit.AgentID INNER JOIN
      CreditAccount ON AgentToCredit.CreditID = CreditAccount.CreditID
WHERE (AgentToCredit.ValidFlag = 'Y') AND (AgentInfo.DealStatus = 'V')

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_AgentTicketOut
AS
SELECT dbo.AgentInfo.AgentID, dbo.AgentInfo.AgentName, dbo.AgentJZ.BookAgentID, 
      dbo.AgentJZ.TicketOutAgentID
FROM dbo.AgentInfo INNER JOIN
      dbo.AgentJZ ON dbo.AgentInfo.AgentID = dbo.AgentJZ.TicketOutAgentID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_Company
AS
SELECT CompanyInfo.CompanyId, CompanyInfo.Description, CompanyLocale.Locale, 
      CompanyLocale.CompanyName
FROM CompanyInfo INNER JOIN
      CompanyLocale ON CompanyInfo.CompanyId = CompanyLocale.CompanyID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_FuncStr
AS
SELECT UserFunction.UserID, Function.Description, Function.OrderNum
FROM UserFunction INNER JOIN
      Function ON UserFunction.FunctionID = Function.FunctionID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_FunctionGroup
AS
SELECT GroupFunction.GroupID, GroupFunction.FunctionID, View_GroupInfo.Description, 
      View_GroupInfo.Locale, View_GroupInfo.GroupName
FROM GroupFunction INNER JOIN
      View_GroupInfo ON GroupFunction.GroupID = View_GroupInfo.GroupID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_FunctionInfo
AS
SELECT Function.FunctionID, Function.Description, Function.Conflict, 
      FunctionLocale.Locale, FunctionLocale.FunctionName, Function.OrderNum
FROM Function LEFT OUTER JOIN
      FunctionLocale ON Function.FunctionID = FunctionLocale.FunctionID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_FunctionUser
AS
SELECT UserFunction.UserID, UserFunction.FunctionID, View_UserInfo.LoginID, 
      View_UserInfo.Password, View_UserInfo.EndDate, View_UserInfo.Name, 
      View_UserInfo.ContactInfo, View_UserInfo.AgentID, View_UserInfo.Sex, 
      View_UserInfo.CompanyID, View_UserInfo.UseObject, View_UserInfo.Status, 
      View_UserInfo.AgentName, View_UserInfo.Locale, View_UserInfo.CompanyName, 
      View_UserInfo.UseObjectName
FROM UserFunction INNER JOIN
      View_UserInfo ON UserFunction.UserID = View_UserInfo.UserID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_GroupFuncStr
AS
SELECT dbo.UserGroup.UserID, dbo.GroupFunction.GroupID, 
      dbo.GroupFunction.FunctionID, dbo.Function.OrderNum, 
      dbo.Function.Description
FROM dbo.GroupFunction INNER JOIN
      dbo.Function ON 
      dbo.GroupFunction.FunctionID = dbo.Function.FunctionID INNER JOIN
      dbo.UserGroup ON dbo.GroupFunction.GroupID = dbo.UserGroup.GroupID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_GroupFunction
AS
SELECT GroupFunction.GroupID, GroupFunction.FunctionID, View_FunctionInfo.Description, 
      View_FunctionInfo.Conflict, View_FunctionInfo.Locale, 
      View_FunctionInfo.FunctionName, View_FunctionInfo.OrderNum
FROM GroupFunction INNER JOIN
      View_FunctionInfo ON GroupFunction.FunctionID = View_FunctionInfo.FunctionID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_GroupInfo
AS
SELECT GroupInfo.GroupID, GroupInfo.Description, GroupLocale.Locale, 
      GroupLocale.GroupName
FROM GroupInfo INNER JOIN
      GroupLocale ON GroupInfo.GroupID = GroupLocale.GroupID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_GroupUser
AS
SELECT UserGroup.UserID, UserGroup.GroupID, View_UserInfo.LoginID, 
      View_UserInfo.Password, View_UserInfo.EndDate, View_UserInfo.Name, 
      View_UserInfo.Sex, View_UserInfo.ContactInfo, View_UserInfo.AgentID, 
      View_UserInfo.CompanyID, View_UserInfo.UseObject, View_UserInfo.Status, 
      View_UserInfo.AgentName, View_UserInfo.Locale, View_UserInfo.CompanyName, 
      View_UserInfo.UseObjectName
FROM UserGroup INNER JOIN
      View_UserInfo ON UserGroup.UserID = View_UserInfo.UserID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_UserFunction
AS
SELECT UserFunction.UserID, UserFunction.FunctionID, View_FunctionInfo.Description, 
      View_FunctionInfo.Conflict, View_FunctionInfo.Locale, 
      View_FunctionInfo.FunctionName, View_FunctionInfo.OrderNum
FROM UserFunction INNER JOIN
      View_FunctionInfo ON UserFunction.FunctionID = View_FunctionInfo.FunctionID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_UserGroup
AS
SELECT UserGroup.UserID, UserGroup.GroupID, View_GroupInfo.Description, 
      View_GroupInfo.Locale, View_GroupInfo.GroupName
FROM UserGroup INNER JOIN
      View_GroupInfo ON UserGroup.GroupID = View_GroupInfo.GroupID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_UserGroupFunction
AS
SELECT View_UserGroup.UserID, View_UserGroup.GroupID, View_UserGroup.Locale, 
      View_UserGroup.GroupName, View_GroupFunction.FunctionID, 
      View_GroupFunction.FunctionName, View_GroupFunction.OrderNum, 
      View_GroupFunction.Description
FROM View_UserGroup INNER JOIN
      View_GroupFunction ON 
      View_UserGroup.GroupID = View_GroupFunction.GroupID AND 
      View_UserGroup.Locale = View_GroupFunction.Locale

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_UserInfo
AS
SELECT dbo.UserInfo.UserID, dbo.UserInfo.LoginID, dbo.UserInfo.Password, 
      dbo.UserInfo.EndDate, dbo.LoginInfo.Name, dbo.LoginInfo.Sex, 
      dbo.LoginInfo.ContactInfo, dbo.LoginInfo.AgentID, dbo.LoginInfo.CompanyID, 
      dbo.LoginInfo.UseObject, dbo.UserInfo.Status, dbo.AgentInfo.AgentName, 
      dbo.UseObject.UseObjectName, dbo.CompanyLocale.CompanyName, 
      dbo.CompanyLocale.Locale
FROM dbo.LoginInfo INNER JOIN
      dbo.UserInfo ON dbo.LoginInfo.LoginID = dbo.UserInfo.LoginID INNER JOIN
      dbo.AgentInfo ON dbo.LoginInfo.AgentID = dbo.AgentInfo.AgentID INNER JOIN
      dbo.UseObject ON dbo.LoginInfo.UseObject = dbo.UseObject.UseObject INNER JOIN
      dbo.CompanyLocale ON dbo.LoginInfo.CompanyID = dbo.CompanyLocale.CompanyID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.View_VolidLogin
AS
SELECT dbo.LoginInfo.LoginID, dbo.LoginInfo.Name, dbo.LoginInfo.Sex, 
      dbo.LoginInfo.AgentID, dbo.LoginInfo.CompanyID, dbo.LoginInfo.ContactInfo, 
      dbo.LoginInfo.UseObject, dbo.UseObject.UseObjectName, 
      dbo.CompanyLocale.Locale, dbo.CompanyLocale.CompanyName, 
      dbo.AgentInfo.AgentName, dbo.AgentInfo.AgentType, dbo.UserHistory.Status, 
      dbo.UserHistory.EndDate, dbo.UserHistory.StartDate
FROM dbo.AgentInfo INNER JOIN
      dbo.LoginInfo ON dbo.AgentInfo.AgentID = dbo.LoginInfo.AgentID INNER JOIN
      dbo.UseObject ON dbo.LoginInfo.UseObject = dbo.UseObject.UseObject INNER JOIN
      dbo.CompanyLocale ON 
      dbo.LoginInfo.CompanyID = dbo.CompanyLocale.CompanyID INNER JOIN
      dbo.UserHistory ON dbo.LoginInfo.LoginID = dbo.UserHistory.LoginID
WHERE (dbo.UserHistory.Status = 'V')

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
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
truncate table AgentNet
insert AgentNet (AgentID,ManageAgentID) select * from #dest order by AgentID

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

