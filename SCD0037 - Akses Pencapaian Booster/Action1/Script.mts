Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dtSidebarMenu, dtSidebar_SubMenu, dtSidebar_Submenu_Submenu, dtMenu_Merchant_Pembelian, jenisPembelian, idPelanggan, Nominal, NoReff, PINTransaksi, dt_UserLogin, dt_Periode
Dim  noRek, noJurnal, trxDate 

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DigisalesLib_Report.xlsx", "SCD0037_Hak Akses Report Pencapaian Booster.xlsx", "SCD0037")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Login Sebagai : " & dt_UserLogin, "Report Periode : " & dt_Periode))

REM ------- Digisales
Call DA_Login()
call FR_GoTo_SidebarMenu(dtSidebarMenu)
Call Search_PencapaianBooster()
Call DA_Logout("0")

Call CreateSessionHeidi_NoSS()
call ExecuteSQL()


'Call DA_Logout("0")

 @@ script infofile_;_ZIP::ssf13.xml_;_

'Call spReportInitiate()

REM ------- Digisales
'Call DA_Login()
'Call FR_GoTo_SidebarMenu(dtSidebarMenu)
'Call FR_GoTo_Sidebar_SubMenu(dtSidebar_SubMenu)
'Call DA_GoTo_Sidebar_Submenu_SubMenu(dtSidebar_Submenu_Submenu)
'Call FR_GoTo_Merchant_MenuPembelian(dtMenu_Merchant_Pembelian)
'
'If dt_TCID = "AGN0016" Then
'	Call FR_Pembelian_PLNPrabayar(jenisPembelian, idPelanggan, PINTransaksi, Nominal)
'	Call FR_Pembelian_PLNPrabayar_TrxSukses(Nominal)
'	Call FR_LaporanTransaksi_GetJurnalDate(dtSidebarMenu)
'	Call spGetDatatable()
'
'	
'ElseIf dt_TCID = "AGN0017" Then	
'	Call FR_Pembelian_PLNPrabayar(jenisPembelian, idPelanggan, PINTransaksi, Nominal)
'	Call FR_Pembelian_PLNPrabayar_IDPelSalah()
'
'ElseIf dt_TCID = "AGN0018" Then
'	Call FR_PLN_ManualAdvice(jenisPembelian, idPelanggan, NoReff, PINTransaksi)
'	Call FR_PLNMA_TransaksiSukses()
'End If


	REM ------ Verifications
'	If dt_TCID = "AGN0016" Then
'		Call spGetDatatable()
'		Call DA_LogCheckpoint(noJurnal, idPelanggan, dt_TCID)
'		Call DA_SQLCheckpoint (noJurnal, trxDate)
'		Call DA_IconsCheckpoint(noRek, noJurnal)
'	ElseIf dt_TCID = "AGN0018" Then
'		Call spGetDatatable()
'		Call DA_LogCheckpoint(noJurnal, idPelanggan, dt_TCID)
'	End If

Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathDigisales, LibReport, LibRepo, objSysInfo
	Dim tempDigisalesPath, tempDigisalesPath2, PathDigisales
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempDigisalesPath 	= Environment.Value("TestDir")
	tempDigisalesPath2 	= InStrRev(tempDigisalesPath, "\")
	PathDigisales 		= Left(tempDigisalesPath, tempDigisalesPath2)
	
	LibPathDigisales		= PathDigisales & "Lib_Repo_Excel\LibDigisales\"
	LibReport			= PathDigisales & "Lib_Repo_Excel\LibReport\"
	LibRepo				= PathDigisales & "Lib_Repo_Excel\Repo\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	rem ---- Digisales lib
	LoadFunctionLibrary (LibPathDigisales & "DigisalesLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDigisales & "Digisales_PencapaianBooster.qfl")
	LoadFunctionLibrary (LibPathDigisales & "Digisales_Heidi.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Akses_Pencapaian_Booster.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Home_Digisales_Web.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Heidi.tsr")
	
'	REM --- Transaction Library
'	LoadFunctionLibrary (LibPathDigisales & "Digisales46_Pembelian.qfl")
'	Call RepositoriesCollection.Add(LibRepo & "RP_Merchant_Pembelian.tsr")
'	
'	REM --- Laporan Transaksi
'	LoadFunctionLibrary (LibPathDigisales & "Digisales46_LaporanTransaksi.qfl")
'	Call RepositoriesCollection.Add(LibRepo & "RP_LaporanTransaksi.tsr")
'	
'	REM --- Verifications Library
'	LoadFunctionLibrary (LibPathDigisales & "DigisalesLib_Verifikasi.qfl")
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_UserLogin				= DataTable.Value("USER",dtLocalSheet)
	dt_Periode					= DataTable.Value("TEXT2",dtLocalSheet) &" "& DataTable.Value("TEXT3",dtLocalSheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
'	
'	REM ---------- Menu
	dtSidebarMenu				= DataTable.Value("SIDEBAR_MENU" ,dtLocalSheet)
'	dtSidebar_toSubmenu			= DataTable.Value("SIDEBAR_SUBMENU" ,dtLocalSheet)
'	dtSidebar_Submenu_Submenu	= DataTable.Value("SIDEBAR_SUBMENU_SUBMENU", dtLocalSheet)
'	dtMenu_Merchant_Pembelian	= DataTable.Value("MENU_MERCHANT_PEMBELIAN" ,dtLocalSheet)
'
'	REM ---- Transaksi
'	jenisPembelian				= DataTable.Value("JENIS_PEMBELIAN_PLN" ,dtLocalSheet)
'	idPelanggan					= DataTable.Value("ID_PELANGGAN" ,dtLocalSheet)
'	NoReff						= DataTable.Value("NO_REFF" ,dtLocalSheet)
'	Nominal						= DataTable.Value("NOMINAL" ,dtLocalSheet)
'	PINTransaksi				= DataTable.Value("PIN_TRX" ,dtLocalSheet)
'	
'	REM ------  Verifications
'	noJurnal 					= DataTable.Value("OUT_NO_JURNAL", dtLocalSheet)
'	trxDate  					= DataTable.Value("OUT_TRX_DATE", dtLocalSheet)
'	noRek 						= DataTable.Value("NO_REKENING", dtlocalsheet)
End Sub
