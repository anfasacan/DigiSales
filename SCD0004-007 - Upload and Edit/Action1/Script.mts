Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dtSidebarMenu, dt_UserLogin
Dim UploadPathDigisales, dt_File

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DigisalesLib_Report.xlsx", "SCD0004-007 - Upload List File and Edit.xlsx", "SCD0004")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Login Sebagai : " & dt_UserLogin, "File Upload : " & dt_File))

REM ------- Digisales
Call DA_Login()
call FR_GoTo_SidebarMenu(dtSidebarMenu)
call UploadFileDistribution()
Call EditFileDistribution()
Call DA_Logout("0")

Call OpenFile( UploadPathDigisales , dt_File, "CHROME" )
Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathDigisales, LibReport, LibRepo, objSysInfo
	Dim tempDigisalesPath, tempDigisalesPath2, PathDigisales
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempDigisalesPath 	= Environment.Value("TestDir")
	tempDigisalesPath2 	= InStrRev(tempDigisalesPath, "\")
	PathDigisales 		= Left(tempDigisalesPath, tempDigisalesPath2)
	
	LibPathDigisales	= PathDigisales & "Lib_Repo_Excel\LibDigisales\"
	LibReport			= PathDigisales & "Lib_Repo_Excel\LibReport\"
	LibRepo				= PathDigisales & "Lib_Repo_Excel\Repo\"
	UploadPathDigisales	= PathDigisales & "File_Upload\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	rem ---- Digisales lib
	LoadFunctionLibrary (LibPathDigisales & "DigisalesLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDigisales & "Digisales_FileDistribution.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_Home_Digisales_Web.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Upload_File_List.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Function.tsr")
	

End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_UserLogin		= DataTable.Value("USER",dtLocalSheet)
	dt_File				= DataTable.Value("FILE1",dtLocalSheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
	REM-------------Sidebar
	dtSidebarMenu = DataTable.Value("SIDEBAR_MENU", dtLocalSheet)

End Sub

