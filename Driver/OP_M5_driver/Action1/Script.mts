Services.StartTransaction "M5_opencart"

mrowcount=datatable.GetSheet("Action1").GetRowcount
msgbox mrowcount
For i = 1 To mrowcount Step 1
	Datatable.SetCurrentRow(i)
	Modexe=Datatable("Moduleexe","Action1")
	msgbox Modexe
	If Modexe="Y" Then
		Modid=Datatable("ModuleID","Action1")
		msgbox Modid
		trowcount=datatable.GetSheet("Action2").GetRowCount
		msgbox trowcount
		For j = 1 To trowcount Step 1
			Datatable.SetCurrentRow(j)
			If Modid=Datatable("ModuleID","Action2") and Datatable("Testcaseexe","Action2")="Y" Then
				testcaseid=Datatable("TestcaseID","Action2")
				msgbox testcaseid
				tsrowcount=datatable.GetSheet("Action3").GetRowCount
		        msgbox tsrowcount
		        For k = 1 To tsrowcount Step 1
			    Datatable.SetCurrentRow(k)
			    If testcaseid=Datatable("TestcaseId","Action3") Then
				keyword=Datatable("Keyword","Action3")
				msgbox keyword
				Select Case (keyword)

					Case "M5_R"
					Case "M5_UMR"
					d1=datatable.GetSheet("Action4").GetRowCount
 					For R = 1 To d1 Step 1
					datatable.SetCurrentRow(R)
					Call Register(datatable("FirstName","Action4"),datatable("LastName","Action4"),datatable("Email","Action4"),datatable("Telephone","Action4"),datatable("Password","Action4"),datatable("ConfirmPassword","Action4"))
					Next
					
					Case "M5_FP"
					Call ForgotPassword()
					
					Case "M5_L"
					drowcount=datatable.GetSheet("Action5").GetRowCount
					For L = 1 To drowcount Step 1
						datatable.SetCurrentRow(L)
					Call Login(datatable("Email","Action5"),datatable("Password","Action5"))
					Next
					
					Case "M5_ED"
					Call EditDetails()
					
				End Select
				
			End If
		Next
	End If

Next
End If
Next

Services.EndTransaction "M5_opencart" @@ script infofile_;_ZIP::ssf43.xml_;_
