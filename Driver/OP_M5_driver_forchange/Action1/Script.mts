Services.StartTransaction "Module5_MyAccount"

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
				
'					Case "M5_URL"
'					Call OpenURL()

					Case "M5_R"
					Case "M5_UMR"
'					Call Register("anki","gawa","asdfghjkl@gmail.com","12345","12345","12345")
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
					'Case "M5_UML"
				'	Call Login(datatable("Email","Action5"),datatable("Password","Action5"))
					
					Case "M5_ED"
					Call EditDetails()
					
				End Select
				
			End If
		Next
	End If

Next
End If
Next

Services.EndTransaction "Module5_MyAccount"




'Browser("Your Store").Page("Your Store").Link("My Account").Click @@ script infofile_;_ZIP::ssf1.xml_;_
'Browser("Your Store").Page("Your Store").Link("Register").Click @@ script infofile_;_ZIP::ssf2.xml_;_
'Browser("Register Account").Page("Register Account").WebEdit("firstname").Set "ANKITA" @@ script infofile_;_ZIP::ssf3.xml_;_
'Browser("Register Account").Page("Register Account").WebEdit("lastname").Set "GAWANDE" @@ script infofile_;_ZIP::ssf4.xml_;_
'Browser("Register Account").Page("Register Account").WebEdit("email").Set "aannkkiittaa2501@gmail.com" @@ script infofile_;_ZIP::ssf5.xml_;_
'Browser("Register Account").Page("Register Account").WebEdit("telephone").Set "231211" @@ script infofile_;_ZIP::ssf6.xml_;_
'Browser("Register Account").Page("Register Account").WebEdit("password").SetSecure "62503b63ae620c3c9507f4c60e666ea8" @@ script infofile_;_ZIP::ssf7.xml_;_
'Browser("Register Account").Page("Register Account").WebEdit("confirm").SetSecure "62503b6a5e01d953be1554c938a860af" @@ script infofile_;_ZIP::ssf8.xml_;_
'Browser("Register Account").Page("Register Account").WebRadioGroup("newsletter").Select "1" @@ script infofile_;_ZIP::ssf9.xml_;_
'Browser("Register Account").Page("Register Account").WebCheckBox("agree").Set "ON" @@ script infofile_;_ZIP::ssf10.xml_;_
'Browser("Register Account").Page("Register Account").WebButton("Continue").Click @@ script infofile_;_ZIP::ssf11.xml_;_
'Browser("Your Account Has Been").Page("Your Account Has Been").WebButton("Continue").Click
'Browser("My Account").Page("My Account").Link("My Account").Click @@ script infofile_;_ZIP::ssf13.xml_;_
'Browser("My Account").Page("My Account").Link("Logout").Click @@ script infofile_;_ZIP::ssf14.xml_;_


'Browser("Account Logout").Page("Account Logout").Link("My Account").Click @@ script infofile_;_ZIP::ssf15.xml_;_
'Browser("Account Logout").Page("Account Logout").Link("Login").Click @@ script infofile_;_ZIP::ssf16.xml_;_
'Browser("Account Login").Page("Account Login").Link("Forgotten Password").Click @@ script infofile_;_ZIP::ssf17.xml_;_
'Browser("Forgot Your Password?").Page("Forgot Your Password?").WebEdit("email").Set "ankigawande01@gmail.com" @@ script infofile_;_ZIP::ssf18.xml_;_
'Browser("Forgot Your Password?").Page("Forgot Your Password?").WebButton("Continue").Click @@ script infofile_;_ZIP::ssf19.xml_;_


'Browser("Account Login").Page("Account Login").Link("My Account").Click @@ script infofile_;_ZIP::ssf20.xml_;_
'Browser("Account Login").Page("Account Login").Link("Login").Click @@ script infofile_;_ZIP::ssf21.xml_;_
'Browser("Account Login").Page("Account Login").WebEdit("email").Set "ankita" @@ script infofile_;_ZIP::ssf22.xml_;_
'Browser("Account Login").Page("Account Login").WebEdit("password").SetSecure "62503bbc1f2ad2727c37" @@ script infofile_;_ZIP::ssf23.xml_;_
'Browser("Account Login").Page("Account Login").WebButton("Login").Click @@ script infofile_;_ZIP::ssf24.xml_;_


'Browser("Account Login").Page("Account Login").Link("My Account").Click @@ script infofile_;_ZIP::ssf25.xml_;_
'Browser("Account Login").Page("Account Login").Link("Login").Click @@ script infofile_;_ZIP::ssf26.xml_;_
'Browser("Account Login").Page("Account Login").WebEdit("email").Set "ankigawande01@gmail.com" @@ script infofile_;_ZIP::ssf27.xml_;_
'Browser("Account Login").Page("Account Login").WebEdit("password").SetSecure "62503bcf8e1a51609c9cdda9b5aed7a4" @@ script infofile_;_ZIP::ssf28.xml_;_
'Browser("Account Login").Page("Account Login").WebButton("Login").Click @@ script infofile_;_ZIP::ssf29.xml_;_


'Browser("My Account").Page("My Account").Link("Edit your account information").Click @@ script infofile_;_ZIP::ssf30.xml_;_
'Browser("My Account Information").Page("My Account Information").Link("Back").Click @@ script infofile_;_ZIP::ssf31.xml_;_
'Browser("My Account").Page("My Account").Link("Change your password").Click @@ script infofile_;_ZIP::ssf32.xml_;_
'Browser("Change Password").Page("Change Password").Link("Back").Click @@ script infofile_;_ZIP::ssf33.xml_;_
'Browser("My Account").Page("My Account").Link("Modify your address book").Click @@ script infofile_;_ZIP::ssf34.xml_;_
'Browser("Address Book").Page("Address Book").Link("Back").Click @@ script infofile_;_ZIP::ssf35.xml_;_
'Browser("My Account").Page("My Account").Link("Modify your wish list").Click @@ script infofile_;_ZIP::ssf36.xml_;_
'Browser("My Wish List").Page("My Wish List").WebButton("Continue").Click @@ script infofile_;_ZIP::ssf37.xml_;_
'Browser("My Account").Page("My Account").Link("My Account").Click @@ script infofile_;_ZIP::ssf38.xml_;_
'Browser("My Account").Page("My Account").Link("Logout").Click @@ script infofile_;_ZIP::ssf39.xml_;_


 @@ script infofile_;_ZIP::ssf40.xml_;_
