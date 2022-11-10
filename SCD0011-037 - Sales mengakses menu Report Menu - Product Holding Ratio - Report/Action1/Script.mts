Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dt_UserLogin, dt_Bulan, dt_Tahun, dt_File1
Dim DownloadPath

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DigisalesLib_Report.xlsx", "SCD0011-037 - Sales mengakses menu Report Menu - Product Holding Ratio - Report.xlsx", "SCD0011")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Login Sebagai : " & dt_UserLogin, "Data Periode : " & dt_Bulan &" "& dt_Tahun, "Nama File : " & dt_File1))

REM ------- Digisales Mobile
Call DA_LoginMobile()
Call GoToSubNavbar()
Call GoToSubSubNavbar()
Call GenerateProductHolding()
Call ListProductHolding()
Call ExportExcel()
Call DA_LogoutMobile("0")

REM ------- Open File EXCEL
Call OpenLastDownloadFile(dt_File1, "EXCEL")

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
	DownloadPath		= "C:\Users\" & objSysInfo.UserName & "\Downloads"
	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	rem ---- Digisales lib
	LoadFunctionLibrary (LibPathDigisales & "DigisalesLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDigisales & "MDigisales_Home.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_MDigisales_Home.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_MDigisales_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_MDigisales_Profile.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Navbar.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Function.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")

End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_UserLogin				= DataTable.Value("USER",dtLocalSheet)
	dt_Bulan					= DataTable.Value("TEXT2",dtLocalSheet)
	dt_Tahun					= DataTable.Value("TEXT3",dtLocalSheet)
	dt_File1					= DataTable.Value("FILE1",dtLocalSheet)

	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
End Sub
