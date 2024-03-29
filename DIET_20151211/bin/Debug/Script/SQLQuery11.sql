IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[PROCON_UPDATEONOFFSTATUS_u]') AND type in (N'P', N'PC'))
DROP PROCEDURE [PROCON_UPDATEONOFFSTATUS_u]
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[PROCON_UPDATEORDERDAYS_u]') AND type in (N'P', N'PC'))
DROP PROCEDURE [PROCON_UPDATEORDERDAYS_u]
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[PROCON_UPDATEPROGRAMDAYS_u]') AND type in (N'P', N'PC'))
DROP PROCEDURE [PROCON_UPDATEPROGRAMDAYS_u]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[PROCON_UPDATEPROGRAMDAYS_u]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [PROCON_UPDATEPROGRAMDAYS_u]
	@DocEntry VarChar(10)
As
BEGIN
	Declare @NoofDays As Integer
	Declare @ConsumeDays As Integer
	
	Declare @ProgramID As VarChar(10)
	
	--Normal Delivery
	Declare @Program As Table (RowID Int,ProgramID VarChar(10))
	
	--Canceled Delivery
	Declare @CProgram As Table (RowID Int,ProgramID VarChar(10))
	
	Insert INTO @Program 
	Select ROW_NUMBER() OVER(ORDER BY T0.ProgramID) As ''RowID'',T0.ProgramID  
	From 
	(
		Select Distinct T0.U_ProgramID As ''ProgramID''
		From [DLN1] T0 JOIN [ODLN] T1 On T0.DocEntry = T1.DocEntry Where T0.DocEntry = @DocEntry
		And ISNULL(T0.U_ProgramID,'''') <> ''''
		And T1.CANCELED = ''N''
	) T0
	
	Insert INTO @CProgram 
	Select ROW_NUMBER() OVER(ORDER BY T0.ProgramID) As ''RowID'',T0.ProgramID  
	From 
	(
		Select Distinct T0.U_ProgramID As ''ProgramID''
		From [DLN1] T0 JOIN [ODLN] T1 On T0.DocEntry = T1.DocEntry Where T0.DocEntry = @DocEntry
		And ISNULL(T0.U_ProgramID,'''') <> ''''
		And T1.CANCELED = ''C''
	) T0
	
		
	Declare @intRow As Int = (Select Count(*) From @Program)	
	While (@intRow > 0)
		BEGIN		
				
			SET @ProgramID = (SELECT ProgramID From @Program Where RowID = @intRow)			
					
			Set @NoofDays = 
			(
				Select Count(U_DelDate) From
				(
					Select Distinct U_DelDate From [DLN1] T0
					JOIN [ODLN] T1 On T0.DocEntry = T1.DocEntry
					Where U_ProgramID = @ProgramID 
					And T1.CANCELED = ''N''
				) T0
			)			
		
			Update T0 Set T0.U_RemDays = (((ISNULL(T0.U_NoOfDays,0)) + (ISNULL(T0.U_FreeDays,0))) - @NoofDays)
			,U_DocStatus = (Case When (T0.U_NoOfDays - @NoofDays) <= 0 Then ''C'' ELSE ''O'' END)
			From [@Z_OCPM] T0 
			Where DocEntry = @ProgramID
			
			Update T0 Set T0.U_DelDays = ISNULL(@NoofDays,0)
			From [@Z_OCPM] T0 
			Where DocEntry = @ProgramID
			
			Set @intRow	= @intRow - 1
				
		END
		
				
		Declare @intRow1 As Int = (Select Count(*) From @CProgram)	
		While (@intRow1 > 0)
		BEGIN									
				
			SET @ProgramID = (SELECT ProgramID From @CProgram Where RowID = @intRow1)	
			
			Set @ConsumeDays = 
			(
				Select Count(U_DelDate) From
				(
					Select Distinct U_DelDate From [DLN1] T0
					JOIN [ODLN] T1 On T0.DocEntry = T1.DocEntry
					Where U_ProgramID = @ProgramID 
					And T1.CANCELED = ''N''
				) T0
			)	
			
			Set @NoofDays = 
			(
				Select Count(U_DelDate) From
				(
					Select Distinct U_DelDate From [DLN1] T0
					JOIN [ODLN] T1 On T0.DocEntry = T1.DocEntry
					Where U_ProgramID = @ProgramID 
					And T1.CANCELED = ''C''
				) T0
			)
			
			Update T0 Set T0.U_RemDays = (ISNULL(T0.U_NoOfDays,0)+ ISNULL(T0.U_FreeDays,0) - ISNULL(@ConsumeDays,0))
			,U_PToDate = (U_PToDate + @NoofDays)
			,U_DocStatus = (Case When (ISNULL(T0.U_NoOfDays,0)+ ISNULL(T0.U_FreeDays,0) - ISNULL(@ConsumeDays,0)) > 0 Then ''O'' ELSE ''C'' END)
			From [@Z_OCPM] T0 
			Where T0.DocEntry = @ProgramID
			
			Update T0 Set T0.U_DelDays =
			(Case 
			When (ISNULL(@ConsumeDays,0) - ISNULL(@NoofDays,0)) > 0 THEN (ISNULL(@ConsumeDays,0) - ISNULL(@NoofDays,0))
			Else 0 
			End)
			From [@Z_OCPM] T0 
			Where DocEntry = @ProgramID
			
			Set @intRow1 = @intRow1 - 1
				
		END
		
END' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[PROCON_UPDATEORDERDAYS_u]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [PROCON_UPDATEORDERDAYS_u]
	@DocEntry VarChar(10)
As
BEGIN
	Declare @NoofDays As Integer
	Declare @ConsumeDays As Integer
	
	Declare @ProgramID As VarChar(10)
	
	--Normal Order
	Declare @Program As Table (RowID Int,ProgramID VarChar(10))
	
	--Canceled Order
	Declare @CProgram As Table (RowID Int,ProgramID VarChar(10))
	
	Insert INTO @Program 
	Select ROW_NUMBER() OVER(ORDER BY T0.ProgramID) As ''RowID'',T0.ProgramID  
	From 
	(
		Select Distinct T0.U_ProgramID As ''ProgramID''
		From [RDR1] T0 JOIN [ORDR] T1 On T0.DocEntry = T1.DocEntry Where T0.DocEntry = @DocEntry
		And ISNULL(T0.U_ProgramID,'''') <> ''''
		And T1.CANCELED = ''N''
	) T0
	
	Insert INTO @CProgram 
	Select ROW_NUMBER() OVER(ORDER BY T0.ProgramID) As ''RowID'',T0.ProgramID  
	From 
	(
		Select Distinct T0.U_ProgramID As ''ProgramID''
		From [RDR1] T0 JOIN [ORDR] T1 On T0.DocEntry = T1.DocEntry Where T0.DocEntry = @DocEntry
		And ISNULL(T0.U_ProgramID,'''') <> ''''
		And T0.TargetType = ''-1'' And T0.LineStatus = ''C''
	) T0
	
	--Normal Order Document	
	Declare @intRow As Int = (Select Count(*) From @Program)	
	While (@intRow > 0)
		BEGIN		
				
			SET @ProgramID = (SELECT ProgramID From @Program Where RowID = @intRow)			
					
			Set @NoofDays = 
			(
				Select Count(U_DelDate) From
				(
					Select Distinct U_DelDate From [RDR1] T0
					JOIN [ORDR] T1 On T0.DocEntry = T1.DocEntry
					Where U_ProgramID = @ProgramID 
					And T1.CANCELED = ''N''
				) T0
			)			
			
			Print @NoofDays
			
			Update T0 Set T0.U_OrdDays = (@NoofDays)			
			From [@Z_OCPM] T0 
			Where DocEntry = @ProgramID
			Set @intRow	= @intRow - 1
				
		END
		
			
		Declare @intRow1 As Int = (Select Count(*) From @CProgram)	
		While (@intRow1 > 0)
		BEGIN									
				
			SET @ProgramID = (SELECT ProgramID From @CProgram Where RowID = @intRow1)	
			
			Set @ConsumeDays = 
			(
				Select Count(U_DelDate) From
				(
					Select Distinct U_DelDate From [RDR1] T0
					JOIN [ORDR] T1 On T0.DocEntry = T1.DocEntry
					Where U_ProgramID = @ProgramID 
					And T1.CANCELED = ''N''
				) T0
			)	
			
			Print @ConsumeDays
			
			Set @NoofDays = 
			(
				
				(Select Count(U_DelDate) From
				(
					Select Distinct U_DelDate From [RDR1] T0
					JOIN [ORDR] T1 On T0.DocEntry = T1.DocEntry
					Where U_ProgramID = @ProgramID 
					And T0.LineStatus = ''C''
					AND T0.TargetType = ''-1''
				) T0) 
			)
			
			Print @NoofDays
			
			Update T0 Set T0.U_OrdDays = 
			(Case 
			When (ISNULL(@ConsumeDays,0) - ISNULL(@NoofDays,0)) > 0 Then (ISNULL(@ConsumeDays,0) - ISNULL(@NoofDays,0))
			Else 0
			End )
			From [@Z_OCPM] T0 
			Where T0.DocEntry = @ProgramID
			
			Set @intRow1 = @intRow1 - 1
				
		END
		
END' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[PROCON_UPDATEONOFFSTATUS_u]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [PROCON_UPDATEONOFFSTATUS_u]
As
Begin

	--Changing Status OFF to ON
	Update T0 SET T0.U_ONOFFSTA = ''F''
	From [@Z_OCPR] T0 
	JOIN
	(
		Select U_CardCode,SUM(ISNULL(U_RemDays,0)) As ''RemDays'' From [@Z_OCPM] 		
		Group By U_CardCode
		Having SUM(ISNULL(U_RemDays,0)) = 0
	) T1 On	T0.U_CardCode = T1.U_CardCode
	Where ISNULL(T0.U_ONOFFSTA,''O'') = ''O''
	
	--Changing Status ON to OFF
	Update T0 SET T0.U_ONOFFSTA = ''O''
	From [@Z_OCPR] T0 
	JOIN
	(
		Select U_CardCode,SUM(ISNULL(U_RemDays,0)) As ''RemDays'' From [@Z_OCPM]
		Group By U_CardCode
		Having SUM(ISNULL(U_RemDays,0)) > 0
	) T1 On	T0.U_CardCode = T1.U_CardCode
	Where ISNULL(T0.U_ONOFFSTA,''O'') = ''F''
	AND T0.U_SuToDt Is Not Null

	--Changing Status ON to OFF - Dont Have Program
	Update T0 SET T0.U_ONOFFSTA = ''F''
	From [@Z_OCPR] T0 
	Where ISNULL(T0.U_ONOFFSTA,''O'') = ''O''
	And T0.U_CardCode Not In (Select Distinct U_CardCode From [@Z_OCPM] Where U_CardCode Is Not Null)
	
End' 
END
GO
