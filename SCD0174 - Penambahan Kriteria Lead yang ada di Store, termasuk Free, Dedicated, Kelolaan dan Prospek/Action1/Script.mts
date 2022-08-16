Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dtSidebarMenu, dtNavbarMenu, dt_UserLogin
Dim UploadPath, dt_File1, iteration

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DigisalesLib_Report.xlsx", "SCD0174 - Penambahan Kriteria Lead yang ada di Store, termasuk Free, Dedicated, Kelolaan dan Prospek.xlsx", "SCD0174")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Login Sebagai : " & dt_UserLogin))

iteration = Environment.Value("ActionIteration")

If iteration = 1 Then
	REM ------- Digisales
	Call DA_Login()
	Call FR_GoTo_SidebarMenu(dtSidebarMenu)
	Call UploadDataLeadsss()
	Call DA_Logout("0")
	
	REM ------ Open File Excel
	Call OpenFile(UploadPath , dt_File1, "EXCEL")
	
ElseIf iteration = 2 Then
	REM ------- Digisales Mobile
	Call DA_LoginMobile()
	Call FR_GoTo_NavbarMenu(dtNavbarMenu)
	Call GoToSubNavbar_Using_TEXT()
	Call DA_LogoutMobile("0")
	
End If

Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathDigisales, LibReport, LibRepo, objSysInfo
	Dim tempDigisalesPath, tempDigisalesPath2, PathDigisales
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempDigisalesPath 	= Environment.Value("TestDir")
	tempDigisalesPath2 	= InStrRev(tempDigisalesPath, "\")
	PathDigisales 		= Left(tempDigisalesPath, tempDigisalesPath2)
	
	UploadPath			= PathDigisales & "File_Upload\"
	LibPathDigisales	= PathDigisales & "Lib_Repo_Excel\LibDigisales\"
	LibReport			= PathDigisales & "Lib_Repo_Excel\LibReport\"
	LibRepo				= PathDigisales & "Lib_Repo_Excel\Repo\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	REM ---- Digisales lib
	
	'Digisales Portal
	LoadFunctionLibrary (LibPathDigisales & "DigisalesLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDigisales & "Digisales_UploadDataLeads.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_Upload_Data_Leads.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Home_Digisales_Web.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Function.tsr")
	
	'Digisales Mobile
	LoadFunctionLibrary (LibPathDigisales & "MDigisales_Store.qfl")
'	LoadFunctionLibrary (LibPathDigisales & "MDigisales_Home.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_MDigisales_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_MDigisales_Store.tsr")
'	Call RepositoriesCollection.Add(LibRepo & "RP_MDigisales_Home.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_MDigisales_Profile.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Navbar.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_UserLogin				= DataTable.Value("USER",dtLocalSheet)
	dt_File1					= DataTable.Value("TEXT1",dtLocalSheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
	REM ---------- Navbar Menu
	dtNavbarMenu				= DataTable.Value("NAVBAR_MENU" ,dtLocalSheet)
	
	REM ---------- Sidebar Menu
	dtSidebarMenu				= DataTable.Value("SIDEBAR_MENU",dtLocalSheet)

End Sub
