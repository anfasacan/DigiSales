Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dtSidebarMenu, dtSidebar_SubMenu, dtSidebar_Submenu_Submenu, dt_UserLogin
Dim dt_File1, dt_Periode, DownloadPath, iteration

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DigisalesLib_Report.xlsx", "SCD0005-003 - Membuka Gembok.xlsx", "SCD0005")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Login Sebagai : " & dt_UserLogin))

REM ------- Digisales

iteration =Environment.Value("ActionIteration")

Call DA_Login()
call FR_GoTo_SidebarMenu(dtSidebarMenu)
 
If ucase(dt_UserLogin) = "SALES" and iteration = 1 Then
	call DownloadDownloadFile()
	Call Get_nama_file
	Call OpenLastDownloadFileNoClose(dt_File1, "EXCEL")
End If

If ucase(dt_UserLogin) = "SALES" and iteration = 3 Then
	Call CekDataTablePadaDownloadFile()
End If

If Ucase(dt_UserLogin) = "ADMIN SLN" Then
	call SearchMonitoringDanReset()
	call MelihatGembokPadaMonitoringDanReset()
	Call MembukaGembokMonitoringDanReset()
	Call MelihatGembokPadaMonitoringDanReset()
End If

Call DA_Logout("0")
Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathDigisales, LibReport, LibRepo, objSysInfo
	Dim tempDigisalesPath, tempDigisalesPath2, PathDigisales
	
	Set objSysInfo 		= Createobject("Wscript.Network")
	Set wshNetwork = CreateObject("WScript.Network")
	
	tempDigisalesPath 	= Environment.Value("TestDir")
	tempDigisalesPath2 	= InStrRev(tempDigisalesPath, "\")
	PathDigisales 		= Left(tempDigisalesPath, tempDigisalesPath2)
	
	LibPathDigisales	= PathDigisales & "Lib_Repo_Excel\LibDigisales\"
	LibReport			= PathDigisales & "Lib_Repo_Excel\LibReport\"
	LibRepo				= PathDigisales & "Lib_Repo_Excel\Repo\"
	DownloadPath		= "C:\Users\" & wshNetwork.UserName & "\Downloads"
'	MsgBox DownloadPath

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	rem ---- Digisales lib
	LoadFunctionLibrary (LibPathDigisales & "DigisalesLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDigisales & "Digisales_FileDistribution.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_Home_Digisales_Web.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Download_File.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Monitoring_Dan_Reset.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Function.tsr")
	
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_UserLogin				= DataTable.Value("USER",dtLocalSheet)
	dt_File1					= DataTable.Value("FILE1",dtLocalSheet)
	dt_Periode					= DataTable.Value("TEXT1",dtLocalSheet)&" - "&DataTable.Value("TEXT2",dtLocalSheet)
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
'	
'	REM ---------- Menu
	dtSidebarMenu				= DataTable.Value("SIDEBAR_MENU" ,dtLocalSheet)
End Sub

