// >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
//
//                                    SLO RESOURCE CODE 
//                         This is the script for the SLO Performance Report
//
//                This script is designed to be called from the main model dialog box
//
//
// >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
//
//
// A few words about step and summary area checkboxes:
// 
//  1. The macro checks for information passed from the calling procedure.  If
//     present, this information is given first priority.
//  2. Check for saved information in a file (PerfSettings.arr in the "_ui"
//     directory. If valid information is found, this information is given
//     second priority.  Upon exiting (without cancel), flag arrays are always
//     saved to this location.
//  3. If no other information is found, the arrays specified in this macro are
//     used.
//     
//  NOTE: The summary table is disabled in this report, as it has not been 
//        adapted to work with this model.

// ****************************************************************************************************************
// This macro sets up some basic information common to many reports
// NOTE: If values here change, many columns will also need to be updated (search fore colhead).
// ****************************************************************************************************************
//      
Dbox "Performance" (Perf)

    init do
        //Use a settings buffer to allow the user to cancel changes
        sets = CopyArray(Perf.Settings)
    enditem //init
    
    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    //  Information about scenario and output file
    text "Scenario: " 1,1,10,1
    text "Scenario Name" 11, 1, 79 framed Variable: Perf.Args.Info.Name
    text "Output: " 1, 3, 10
    text "Report filename" 11, 3, 79, 1  framed variable: Perf.File
    
// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Checkboxes for reports

    //********BASIC REPORTS
    //frame "Basic Reports" 1, 5, 90, 19 Prompt: "Basic Reports:"
	frame "Basic Reports" 1, 5, 90, 21 Prompt: "Basic Reports:"

	checkbox "Title Page" 3, 7, 35, 1 
        Variable: sets.Report.[RPT Title Page]
    checkbox "Input Files and Parameters" 3, 8.5, 35, 1 
        Variable: sets.Report.[RPT Files]
    checkbox "Input Network Summary*" 3, 10, 35, 1 
        Variable: sets.Report.[RPT Network Summary]
    checkbox "Land Use Data Summary" 3, 11.5, 35, 1 
        Variable: sets.Report.[RPT LU Data]
    
    //Select all or none in this category
    button "All" 3, 14, 5, 1 do
        Perf.SetAllReports(True, sets, "basic")
    enditem
	
    button "None" 10, 14, 5, 1 do
        Perf.SetAllReports(False, sets, "basic")
    enditem
    
    //********PERFORMANCE REPORTS
    //frame "Performance Reports" 45.8, 5, 45.2, 19  Prompt: "Performance Reports:"
	frame "Performance Reports" 45.8, 5, 45.2, 21  
        Prompt: "Performance Reports:"
    checkbox "Trip Generation Summary" 48, 7, 35, 1 
        Variable: sets.Report.[RPT Trip Generation]
    checkbox "Trip Distribution Summary" same, 8.5, 35, 1 
        Variable: sets.Report.[RPT Trip Distribution]
    checkbox "Trip Length Frequencies" same, 10, 35, 1 
        Variable: sets.Report.[RPT Trip Length Frequencies]
    checkbox "Mode Choice Summary" same, 11.5, 35, 1 
        Variable: sets.Report.[RPT Mode Choice]
	checkbox "Assigned Trips Summary" same, 13, 35, 1 
        Variable: sets.Report.[RPT Assigned Vehicle Trips]
	checkbox "Transit Assignment Summary" same, 14.5, 35, 1 
        Variable: sets.Report.[RPT Transit Assignment]
    checkbox "Daily Assignment Summary*" same, 16, 35, 1 
        Variable: sets.Report.[RPT Vehicle Assignment]
	
    //Select all or none in this category
	button "All" 48, 23.5, 5, 1 do
        Perf.SetAllReports(True, sets, "perf")
    enditem
	
	button "None" 55, 23.5, 5, 1 do
        Perf.SetAllReports(False, sets, "perf")
    enditem
    
    //********VALIDATION REPORTS
	frame "Validation Reports" 1, 16, 45, 10 Prompt: "Validation Reports:"

    checkbox "Daily Validation Summary" 3, 18, 35, 1 
        Variable: sets.Report.[RPT Vehicle Validation DY]
    checkbox "AM Validation Summary*" same, 19.5, 35, 1 
        Variable: sets.Report.[RPT Vehicle Validation AM]
    checkbox "PM Validation Summary*" same, 21, 35, 1 
        Variable: sets.Report.[RPT Vehicle Validation PM]


    //Select all or none in this category
    button "All" 3, 24.5, 5, 1 do
        Perf.SetAllReports(True, sets, "validation")
    enditem
	
    button "None" 10, 24.5, 5, 1 do
        Perf.SetAllReports(False, sets, "validation")
    enditem
    
    //*********SUMMARY AREA SELECTION
	frame "Create Reports For" 1, 27, 45, 11 Prompt: "* Create Reports For:"
    
	checkbox "Entire Model" 3, 29,   15, 1 
        Variable: sets.SumArea.[Entire Model]
	checkbox "SB County" 3, 30.5, 15, 1 
        Variable: sets.SumArea.[SB County]
	checkbox "Custom 1" 3, 32, 15, 1 
        Variable: sets.SumArea.[Custom 1]
	checkbox "Custom 2" 3, 33.5,   15, 1  
        Variable: sets.SumArea.[Custom 2]
        
// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
// Global Buttons

    //*********REPORT SELECTION
	frame "Global Selection" 45.8, 27, 45.2, 8 Prompt: "Global Selection:"
	frame "Global Separator" 45.8, 27, 45.2, 4 Prompt: "Global Selection:"
	button "Select All Reports" 50, 28.8, 17, 1.5 do
        Perf.SetAllReports(True, sets)
    enditem
    
	button "Select No Reports" 70, 28.8, 17, 1.5 do
        Perf.SetAllReports(False, sets)
    enditem
    
	button "Select All Areas" 50, 32.3, 17, 1.5 do
        Perf.SetAllAreas(True, sets)
    enditem
    
	button "Select No Areas" 70, 32.3, 17, 1.5 do
        Perf.SetAllAreas(False, sets)
    enditem

    //**********OK OR CANCEL
    //button "Go" 58, 34, 15, 2 default Prompt: OkText do
	button "Go" 58, 36, 15, 2 default Prompt: "OK" do
        Perf.Settings = CopyArray(sets)     //Commit changes to the Perf object
        SaveArray(sets, Perf.SavedSettings) //Save settings for next time
        Return(True)
    enditem
    
    button "Cancel" 75, same, 15, 2 cancel do
        Return()
    enditem

EndDbox

Class "Performance" (Args) //StartClass

    init do
    //StartMethod
        shared UT

        t = SplitPath(GetInterface())
        ui_dir = t[1] + t[2]
    
        //Identify passed scenario
        self.Args = Args
		expand = self.Args.Param.INI.ExpandSetting.Value

        //FT values
        self.Info.FT = {{"Freeway",              					1},   //Note - Freeway MUST be first to produce a correct screenline report.
                        {"HOV",										2},
						{"Expressway/Parkway",						3},
						{"Principal Arterial",						4},
                        {"Minor Arterial",       					5},
                        {"Major Collector",            				6},
						{"Minor Collector",							7},
                        {"Ramps",       							8},
						{"Truck Lanes",								9},
			            {"Centroid Connectors",						10}}
			           // {"Transit Local",							999}}
        
        //AT values
        self.Info.AT = {{"Core", 						1},
						{"Central Business District",	2},
						{"Urban Business District", 	3},
                        {"Urban", 						4},
                        {"Suburban", 					5},
                        {"Rural", 						6},
			            {"Mountain", 					7}}

        //Default format settings
        self.Formats = null
        self.Formats.Filetype = "html" //html or excel - overridden by filename
		
        //Global Formats
        self.Formats.NumberFormats = "*,."
		
        //html formats
        self.Formats.TablesPerPage = 2 //!!! Remove tables per page
        self.Formats.html.CSSFile = ui_dir + "Style.css"
        self.Formats.html.ScriptFile = ui_dir + "ReportScript.txt"
        
        //Default chart settings
        self.ChartCount = 0 //Chart count for re-drawing in pin area
        self.chartDefaults = null
        self.chartDefaults.Type = "bar"
        self.chartDefaults.Width = 600
        self.chartDefaults.Height = 300
        self.chartDefaults.Colors = {{91, 155, 213}, {237, 125, 49}, {255, 192, 0}, {68, 114, 196}, {112, 173, 71}}
        //Repeat colors to allow large sets 
        orig = CopyArray(self.chartDefaults.Colors)
        for ii = 1 to 5 do
            self.chartDefaults.Colors = self.chartDefaults.Colors + CopyArray(orig)
        end
		
        //excel formats
        h1 = null
        h1.Font.Size = 14
        h1.Font.Bold = True
        
        h2 = null
        h2.Font.Size = 12
        h2.Font.Italic = True

        bluetext = null
        bluetext.Font.Size = 10
        bluetext.Font.Color = self.ExcelRGB(54, 96, 146)
        
        greytext = null
        greytext.Font.Size = 10
        greytext.Font.Color = self.ExcelRGB(54, 96, 146)  
        
        value = null
        value.Font.Size = 10
        value.Font.Color = self.ExcelHEX('000000')
        value.Borders.LineStyle = 1 //xlContinous
        value.Borders.Weight = 2 //xlThin
        value.Borders.ThemeColor = 2 //dark grey
        
        bold = CopyArray(value)
        bold.Font.Bold = True
        
        self.Formats.excel.Styles.h1 = h1
        self.Formats.excel.Styles.h2 = h2
        self.Formats.excel.Styles.bluetext = bluetext
        self.Formats.excel.Styles.greytext = greytext
        self.Formats.excel.Styles.value = value
        self.Formats.excel.Styles.bold = bold
        
        self.Formats.excel.LabelWidth = 150  //in pixels, same as html to be consistent if passed
        self.Formats.excel.DataWidth = 75  //by the macros themselves
        self.Formats.excel.TablesPerPage = 2
        
        self.ExcelRow = 1
        self.ExcelCol = 1 //to keep track of row and columns while writing tables
        
        //Define summary areas
        self.SumArea = null
        
        self.SumArea.[Entire Model].Query = "Select * where ID >= 0"	//SCAG
        self.SumArea.[Entire Model].Network = True
        self.SumArea.[Entire Model].Zones = True
        self.SumArea.[Entire Model].Active = False
        
        self.SumArea.[SB County].Query = "Select * where County = 5"
        self.SumArea.[SB County].Network = True
        self.SumArea.[SB County].Zones = True
        self.SumArea.[SB County].Active = False
        
        self.SumArea.[Custom 1].Query = "Select * Where CUSTOM1 = 1"
        self.SumArea.[Custom 1].Network = True
        self.SumArea.[Custom 1].Zones = True
        self.SumArea.[Custom 1].Active = False
        
        self.SumArea.[Custom 2].Query = "Select * Where CUSTOM2 = 1"
        self.SumArea.[Custom 2].Network = True
        self.SumArea.[Custom 2].Zones = True
        self.SumArea.[Custom 2].Active = False
        
        //Define reports
        self.Report = null
        self.Report.[RPT Title Page].Contents = "Title Page"  //Name in the table of contents
        self.Report.[RPT Title Page].Anchor = "TitlePage"     //HTML anchor (for linking)
        self.Report.[RPT Title Page].Active = 1               //Default status (overriden by settings)
        self.Report.[RPT Title Page].Section = 0              //Report section number
        self.Report.[RPT Title Page].Group = "basic"          //Report group (for on/off buttons)
                     
        self.Report.[RPT Files].Contents = "Files and Settings"
        self.Report.[RPT Files].Anchor = "Files"
        self.Report.[RPT Files].Active = 1
        self.Report.[RPT Files].Section = 1
        self.Report.[RPT Files].Group = "basic"
                     
        self.Report.[RPT Network Summary].Contents = "Input Network Summary"
        self.Report.[RPT Network Summary].Anchor = "Network"
        self.Report.[RPT Network Summary].Active = 1
        self.Report.[RPT Network Summary].Section = 2
        self.Report.[RPT Network Summary].Group = "basic"
                     
        self.Report.[RPT LU Data].Contents = "Land Use Data Summary"
        self.Report.[RPT LU Data].Anchor = "Sociodata"
        self.Report.[RPT LU Data].Active = 1
        self.Report.[RPT LU Data].Section = 3
        self.Report.[RPT LU Data].Group = "basic"
                     
        self.Report.[RPT Trip Generation].Contents = "Trip Generation Summary"
        self.Report.[RPT Trip Generation].Anchor = "Tripgen"
        self.Report.[RPT Trip Generation].Active = 1
        self.Report.[RPT Trip Generation].Section = 4
        self.Report.[RPT Trip Generation].Group = "perf"
                     
        self.Report.[RPT Trip Distribution].Contents = "Trip Distribution Summary"
        self.Report.[RPT Trip Distribution].Anchor = "Tripdist"
        self.Report.[RPT Trip Distribution].Active = 1
        self.Report.[RPT Trip Distribution].Section = 5
        self.Report.[RPT Trip Distribution].Group = "perf"
        
        self.Report.[RPT Trip Length Frequencies].Contents = "Trip Length Frequencies"
        self.Report.[RPT Trip Length Frequencies].Anchor = "TLFD"
        self.Report.[RPT Trip Length Frequencies].Active = 1
        self.Report.[RPT Trip Length Frequencies].Section = 4
        self.Report.[RPT Trip Length Frequencies].Group = "perf"
        
        self.Report.[RPT Mode Choice].Contents = "Mode Choice Summary"
        self.Report.[RPT Mode Choice].Anchor = "Mode"
        self.Report.[RPT Mode Choice].Active = 1
        self.Report.[RPT Mode Choice].Section = 6
        self.Report.[RPT Mode Choice].Group = "perf"
        
        self.Report.[RPT Assigned Vehicle Trips].Contents = "Assigned Vehicle Trip Summary"
        self.Report.[RPT Assigned Vehicle Trips].Anchor = "Assign"
        self.Report.[RPT Assigned Vehicle Trips].Active = 1
        self.Report.[RPT Assigned Vehicle Trips].Section = 7
        self.Report.[RPT Assigned Vehicle Trips].Group = "perf"
        
        self.Report.[RPT Transit Assignment].Contents = "Transit Assignment Summary"
        self.Report.[RPT Transit Assignment].Anchor = "Trans"
        self.Report.[RPT Transit Assignment].Active = 1
        self.Report.[RPT Transit Assignment].Section = 8
        self.Report.[RPT Transit Assignment].Group = "perf"
        
        self.Report.[RPT Vehicle Assignment].Contents = "Vehicle Assignment Summary"
        self.Report.[RPT Vehicle Assignment].Anchor = "Assignment"
        self.Report.[RPT Vehicle Assignment].Active = 1
        self.Report.[RPT Vehicle Assignment].Section = 9
        self.Report.[RPT Vehicle Assignment].Group = "perf"
        
        self.Report.[RPT Vehicle Validation DY].Contents = "Daily Vehicle Validation Summary"
        self.Report.[RPT Vehicle Validation DY].Anchor = "Validation_DY"
        self.Report.[RPT Vehicle Validation DY].Active = 1
        self.Report.[RPT Vehicle Validation DY].Section = 10
        self.Report.[RPT Vehicle Validation DY].Group = "validation"
        
        self.Report.[RPT Vehicle Validation AM].Contents = "AM Vehicle Validation Summary"
        self.Report.[RPT Vehicle Validation AM].Anchor = "Validation_AM"
        self.Report.[RPT Vehicle Validation AM].Active = 1
        self.Report.[RPT Vehicle Validation AM].Section = 11
        self.Report.[RPT Vehicle Validation AM].Group = "validation"
        
        self.Report.[RPT Vehicle Validation PM].Contents = "PM Vehicle Validation Summary"
        self.Report.[RPT Vehicle Validation PM].Anchor = "Validation_PM"
        self.Report.[RPT Vehicle Validation PM].Active = 1
        self.Report.[RPT Vehicle Validation PM].Section = 12
        self.Report.[RPT Vehicle Validation PM].Group = "validation"
		       
        
        //Identify report filename
		self.File = self.Args.Info.ModelDir + "Reports\\Performance_Report.html"
        
        // ---------------------------------------------------------------------
        // the remainder should rarely be changed
		
        //Identify report filetype
        t = SplitPath(self.File)
        if t[4] = '.html' or t[4] = '.htm' then 
            self.Formats.Filetype = 'html'
        else if t[4] = '.xls' or t[4] = '.xlsx' or t[4] = '.xlsm' then
            self.Formats.Filetype = 'excel'
        
        //Create default settings
        //1-Attempt to read from default settings file
        self.SavedSettings = ui_dir + "PerfSettings.arr"
        if GetFileInfo(self.SavedSettings) != null then 
            sets = LoadArray(self.SavedSettings)
        else 
            sets = null

        //2-read from Report settings above
        //Settings in the file override the defaults
        self.Settings = null
        for i = 1 to self.Report.Length do
            Key = self.Report[i][1] //looping through Report.*
            Val = self.Report[i][2]
            if FindOption(sets, "Report") != null and FindOption(sets.Report, Key) != null then 
                self.Settings.Report.(Key) = sets.Report.(Key)
            else
                self.Settings.Report.(Key) = Val.Active
                
            Key = null  Val = null //Set to null to prevent array confusion
        end
        for i = 1 to self.SumArea.Length do
            Key = self.SumArea[i][1] //looping through SumArea.*
            Val = self.SumArea[i][2]
            
            if FindOption(sets, "SumArea") != null and FindOption(sets.SumArea, Key) != null then 
                self.Settings.SumArea.(Key) = sets.SumArea.(Key)
            else
                self.Settings.SumArea.(Key) = 1 // TODO: use proper options array --> Val.Active
            
            Key = null  Val = null //Set to null to prevent array confusion
        end
        
        //Default to page 1, first table
        self.page = 1
        self.table = 1
        
    enditem //init
    //EndMethod
    
    //Update the Perf object to use a new set of scenario Args
    Macro "SetArgs" (Args) do
        self.Args = Args
    EndItem
    //EndMethod
    
    //Get scenario settings from a dialog box
    Macro "GetSettings" do
        Return(RunDbox("Performance", self))
    enditem
    //EndMethod
    
    //Activate or de-activate all reports
    Macro "SetAllReports" (val, sets, group) do
    //val = True for all on, False for all off
    //sets = Settings array to use, or null to change performance report object
    //group = null for all reports, group for only reports in a specific group
    //
    //NOTE: If the sets array is passed, the passed array is modified directly
        
        //Use passed settings, or reference Perf.Settings
        if sets = null then sets = self.Settings
        status = if val then 1 else 0
        for i = 1 to sets.Report.Length do
            k = sets.Report[i][1]
            if (group = null or Lower(self.Report.(k).Group) = Lower(group)) and  Lower(self.Report.(k).Group) != 'title' then 
                sets.Report[i][2] = status
        end
    enditem
    //EndMethod
    
    //Activate or de-activate all areas
    Macro "SetAllAreas" (val, sets) do
    //val = True for all on, False for all off
    //sets = Settings array to use, or null to change performance report object
    //
    //NOTE: If the sets array is passed, the passed array is modified directly
    
        //Use passed settings, or reference Perf.Settings
        if sets = null then sets = self.Settings
        status = if val then 1 else 0
        for i = 1 to sets.SumArea.Length do
            sets.SumArea[i][2] = status
        end
    enditem
    //EndMethod

    Macro "CreateReport" do
    
        if !RunMacro("G30 File Close All") then do
            ShowMessage("Performance Report Cancelled")
            Return()
        end
    
        HideDbox()
        
        //Set up the progress bar
        self.canned = null
        progtot = 0
        prognum = 0
        for i = 1 to self.Settings.Report.length do
            Val = self.Settings.Report[i][2]
            if Val then progtot = progtot + 1
            Val = null
        end
        EnableProgressBar("Processing...", 3)
        self.canned = CreateProgressBar("Initializing Report", "False")
        if self.canned then do
            DestroyProgressBar()
            Return()
        end
        
        //Write HTML page headers
        if self.Formats.Filetype = "html" then do
            fp = OpenFile(self.File, "w")
            self.fp = fp
            WriteLine(fp,"<!DOCTYPE html>")
            WriteLine(fp,"<html>")
            WriteLine(fp,"<head>")
            WriteLine(fp,"<title>Model Summary Report</title>")
            WriteLine(fp,'<meta charset="UTF-8">')
            WriteLine(fp, '<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>')
            WriteLine(fp, '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.5.0/Chart.min.js"></script>')
            WriteLine(fp, '<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/themes/smoothness/jquery-ui.css">')
            WriteLine(fp, '<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/jquery-ui.min.js"></script>')
            
            //!!! Keep charts from resizing
            WriteLine(fp, '<script>')
            WriteLine(fp, '  Chart.defaults.global.responsive = false;')
            WriteLine(fp, '</script>')
            WriteLine(fp, "<style type=\"text/css\">")
        
            //Write the CSS Styles
            css_fp = OpenFile(self.Formats.html.CSSFile, "r")
            css_lines = ReadArray(css_fp)
            CloseFile(css_fp)
            for i = 1 to css_lines.length do
                WriteLine(fp, css_lines[i])
            end
            
            WriteLine(fp,"</style>") 
            
            WriteLine(fp,"</head>")
            WriteLine(fp,"<body>")

            //Write the CSS Styles
            jq_fp = OpenFile(self.Formats.html.ScriptFile, "r")
            jq_lines = ReadArray(jq_fp)
            CloseFile(jq_fp)
            for i = 1 to jq_lines.length do
                WriteLine(fp, jq_lines[i])
            end
            
        end
        
        else if self.Formats.Filetype = "excel" then do
        
            //Make sure that an Excel file has the correct extension
            t = SplitPath(self.File)
            self.File = t[1] + t[2] + t[3] + ".xlsx"
            
            //OVERWRITE a pre-existing report - the user is not asked, 
            //   so verify overwrite before calling create
            if GetFileInfo(self.File) != null then do
            
                on error do
                    Opts = null
                    Opts.Buttons = "RetryCancel"
                    ans = MessageBox("Cannot write to Summary report file - Maybe it is open?", Opts)
                    if ans = "Retry" then goto DeleteOldSummary
                    //ShowMessage("Cannot write to Summary report file - Close the file and try again")
                    DestroyProgressBar()
                    DisableProgressBar()
                    Return()
                end
                DeleteOldSummary:
                DeleteFile(self.File)
                on error default
            end
            
        
            //Create the Excel COM object
            XL = CreateCOMObject("Excel.Application")
            self.XL = XL
            
            NewBook = XL.Workbooks.Add()
            XL.WindowState = -4140 //xlMinimized
            XL.Visible = True
            //XL.ScreenUpdating = False
            NewBook.Title = "Summary Report"
            NewBook.SaveAs(self.File)
            //NewBook.SaveAs('C:\\HCAOG Model\\Outputs\\Summary.xlsx')
            
            //Identify initial sheets
            dim InitialSheets[XL.ActiveWorkbook.Sheets.Count]
            for i = 1 to XL.ActiveWorkbook.Sheets.Count do
                SHEET = XL.ActiveWorkbook.Sheets[i]
                InitialSheets[i] = SHEET.Name
            end
            
        end //excel
        else do
            Throw("Invalid report file format specified")
        end
        
        //Write selected reports
        for i = 1 to self.Report.length do
            self.CurrentReportName = self.Report[i][1]  //Name of the report
            self.CurrentReport = self.Report[i][2]  //Report options
            if self.Settings.Report.(self.CurrentReportName) then do
                self.canned = UpdateProgressBar(self.Report.(self.CurrentReportName).Contents, r2i(round(prognum/progtot * 100, 0)))
                if self.canned then do
                    DestroyProgressBar()
                    Return()
                end
                prognum = prognum + 1
            
                RunMacro(self.CurrentReportName, self)
                
                //Reset to first page
                self.page = 1
                self.table = 1
                
                self.ExcelRow = 1
                self.ExcelCol = 1
                
                if self.canned then do
                    DestroyProgressBar()
                    Return()
                end
                
            end
        end
        
        //Finish document
        if Lower(self.Formats.Filetype) = "html" then do
            WriteLine(fp, "</body></html>")
            self.fp = null
            CloseFile(fp)
            fp = null
        end
        else if Lower(self.Formats.Filetype) = "excel" then do
        
            //Delete initial sheets
            for _del = 1 to InitialSheets.length do
                del = InitialSheets[_del]
                for i = 1 to XL.ActiveWorkbook.Sheets.Count do
                    SHEET = XL.ActiveWorkbook.Sheets[i]
                    if del = SHEET.Name then SHEET.Delete()
                end
            end
            
            
            //Activate the first sheet, update the screen, then save and quit
            XL.ActiveWorkbook.Sheets[1].Activate()
            XL.ScreenUpdating = True
            XL.ActiveWorkbook.Save()
            XL.Quit()
            XL = null
            self.XL = null
        
        end
        else do
            Throw("Invalid report file format specified.")
        end

        DestroyProgressBar()
        DisableProgressBar()
        Return(1)
    enditem
    //EndMethod
    
    //Return a list of {name, query} for active summary areas
    Macro "ActiveAreas" (type) do
    //Method must be "Network" or "Zones"
        r = null
        for i = 1 to self.SumArea.length do
            Key = self.SumArea[i][1]
            Val = self.SumArea[i][2]
            if self.Settings.SumArea.(Key) and Val.(type) then do
                r = r + {{Key, Val.Query}}
            end
        end
        Return(r)
    enditem
    //EndMethod
    
    //Returns an array of 2D arrays of cross-classified vector data
    Macro "CrossTab" (FTv, ATv, V, DoMarginals, InOpts) do
    //FTv = Facility Type vector
    //ATv = Area Type vector
    //V = Vector of values
    //DoMarginals = if True, marginals will be added to the result
    //Opts
    //  array .RowList: List of values to include for rows (defaults to FT values)
    //  array .ColList: List of values to include for columns (defaults to AT values)
    //
        //Identify balanced PA file
        bal_file = self.Args.Output.TGN.PAbal.Value
        
        //Process Option to retrieve unique row values, or use default (FT)
        fts = null
        if InOpts.RowList != null and TypeOf(InOpts.RowList) = 'array' then do
            dim fts[InOpts.RowList.Length, 2]
            for i = 1 to fts.length do
                fts[i][1] = InOpts.RowList[i]
                fts[i][2] = i
            end
        end
        else fts = self.Info.FT
        
        //Process Option to retrieve unique col values, or use default (AT)
        ats = null
        if InOpts.ColList != null and TypeOf(InOpts.ColList) = 'array' then do
            dim ats[InOpts.ColList.Length, 2]
            for i = 1 to ats.length do
                ats[i][1] = InOpts.ColList[i]
                ats[i][2] = i
            end
        end
        else ats = self.Info.AT
    
        dim r[fts.length, ats.length]
        for i = 1 to fts.length do
            for j = 1 to ats.length do
                ft = fts[i][2]
                at = ats[j][2]
                
                F = if FTv = ft and ATv = at then V else 0
                r[i][j] = VectorStatistic(F, "Sum", )
            end
        end
        
        if DoMarginals then
            r = self.Marginals(r)
            
        Return(r)
    enditem
    //EndMethod
    
    Macro "Marginals" (array) do
        a = CopyArray(array) //Don't modify the input
        n = a.length
        m = a[1].length

        //Get row totals
        for i = 1 to n do
            if a[i].length <> m then 
                Throw("Non-rectangular array - cannot compute marginals")
            s = 0
            for j = 1 to m do
                s = s + nz(a[i][j])
            end
            a[i] = a[i] + {s}
        end
        //Get column totals
        dim s[m+1]
        for i = 1 to n do
            for j = 1 to m+1 do
                s[j] = nz(s[j]) + nz(a[i][j])
            end
        end
        a = a + {s}
        
        //Return updated array
        return(a)
    enditem
    //EndMethod

    Macro "PageHeader" do
    
        name = self.CurrentReport.Contents
        section = self.CurrentReport.Section
        scen_name = self.Args.Info.Name
        
        fileformat = self.Formats.Filetype
        
        //html format
        if fileformat = "html" then do
            fp = self.fp
            WriteLine(fp,'<h2 id="folder_' + self.CurrentReport.Anchor + '"> <span class="folder" cursor="pointer">&#x25b8; </span>' + Trim(Substitute(self.CurrentReportName, "RPT", "", )) + '</h2>')
            WriteLine(fp,'<a name = "' + self.CurrentReport.Anchor + '"></a>')
            WriteLine(fp,'<div id="data_' + self.CurrentReport.Anchor + '" class="indent_h2">')
        end //html
        
        //excel format
        else if fileformat = "excel" then do
            XL = self.XL
            sheets = XL.ActiveWorkbook.Sheets
            
            //Add a new tab and header if it is the first page
            if self.page = 1 then do
           
                //Add the new sheet at the end of the workbook
                AfterSheet = XL.ActiveWorkbook.Sheets[XL.ActiveWorkbook.Sheets.Count]
                sheets.Add(null, AfterSheet).Name = name
                XL.ActiveSheet.PageSetup.RightHeader = '&"-,Bold Italic"&14Wichita Area Travel Model - ' + 
                                                        scen_name + Char(10) + '&"-,Regular"&11' +
                                                        self.CurrentReportName + ' - Page ' + i2s(section)+
                                                        "."+ '&P'
            end //if first page then add the tab
            
        end
        
    enditem
    //EndMethod
    
    //Write tables (and now charts) to the currently open report file
    Macro "WriteTables" (TableOptsArray, InOpts) do
    //TableOptsArray is an array of options arrays, with one Opts array for
    //  each table to write:
    //   .Name = Name of Table or chart
    //   .Section1 = Sub section 1 name, or null for no sub section
    //   .Section2 = Sub section 2 name, or null for no sub section
    //   .Footnote = footnote to place at end of table or chart
    //   -- only one of the following should be specified --
    //   .Table = Table options array as specified in WriteTable.
    //   .Chart = Chart options array as specified in DrawChart.
    //
    //Other options (InOpts)
    //  .NoHeader = T/F, skip writing the page header if True
    
        TOA = CopyArray(TableOptsArray) //Use a copy of the input opts
        fp = self.fp //quick reference
        
        //Write a Page Header
        if !TOA.NoHeader then self.PageHeader()
        
        //Write tables, with section/sub-section logic
        s1_name = null
        s1_anchor = self.CurrentReport.Anchor + "_MAIN"
        s2_name = null
        for OP in TOA do
        
            //Sub-header 1
            if OP.Section1 != s1_name then do
                //end existing subheader sections if needed
                if s2_name != null then do
                    WriteLine(fp, '</div>')
                    s2_name = null
                end
                if s1_name != null then do
                    WriteLine(fp, "</div>")
                    s2_name = null
                end
                
                //Start new Section 2 (h3)
                if OP.Section1 != null then do
                    s1_name = OP.Section1
                    s1_anchor = self.CurrentReport.Anchor + "_" + Substitute(s1_name, " ", "_", )
                    WriteLine(fp, '<h3 id="folder_'+s1_anchor+'"><span class="folder" cursor="pointer">&#x25b8; </span>')
                    WriteLine(fp, '<a class="clipbd_all"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24">')
                    WriteLine(fp, '<path path="#6D6D6D" fill="#6D6D6D" d="M19 2c-1.229 0-2.18-1.084-3-2h-8c-.82.916-1.771 2-3 2h-3v22h20v-22h-3zm-7 0c.552 0 1 .448 1 1s-.448 1-1 1-1-.448-1-1 .448-1 1-1zm8 20h-3.824c1.377-1.103 2.751-2.51')
                    WriteLine(fp, '3.824-3.865v3.865zm0-8.457c0 4.107-6 2.457-6 2.457s1.518 6-2.638 6h-7.362v-18h4l2.102 2h3.898l2-2h4v9.543z"/></svg></a>')
                    WriteLine(fp, s1_name+'</h3>')
                    WriteLine(fp, '<div class="indent_h3" id="data_'+s1_anchor+'">')
                end else s1_anchor = self.CurrentReport.Anchor + "_MAIN"
            end //subheader 1
            
            //Sub-header 2
            if OP.Section2 != s2_name then do
                
                //end existing subheader 2 sections if needed
                if s2_name != null then do
                    WriteLine(fp, '</div>')
                    s2_name = null
                end
                
                //Start new h4 section
                if OP.Section2 != null then do
                    s2_name = OP.Section2
                    s2_anchor = s1_anchor + "_" + Substitute(s2_name, " ", "_", )
                    WriteLine(fp, '<h4 id="folder_'+s2_anchor+'"><span class="folder" cursor="pointer">&#x25b8; </span>')
                    WriteLine(fp, '<a class="clipbd_all"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24">')
                    WriteLine(fp, '<path path="#6D6D6D" fill="#6D6D6D" d="M19 2c-1.229 0-2.18-1.084-3-2h-8c-.82.916-1.771 2-3 2h-3v22h20v-22h-3zm-7 0c.552 0 1 .448 1 1s-.448 1-1 1-1-.448-1-1 .448-1 1-1zm8 20h-3.824c1.377-1.103 2.751-2.51')
                    WriteLine(fp, '3.824-3.865v3.865zm0-8.457c0 4.107-6 2.457-6 2.457s1.518 6-2.638 6h-7.362v-18h4l2.102 2h3.898l2-2h4v9.543z"/></svg></a>')
                    WriteLine(fp, s2_name+'</h4>')
                    WriteLine(fp, '<div class="indent_h4" id="data_'+s2_anchor+'">')
                end
            end //subheader 2
            

            //Write the table or chart
            WriteLine(fp, "<div>")
            WriteLine(fp, '<h5><span class="pin">&#x1F4CC;&#xFE0E;&nbsp;</span>')
            WriteLine(fp, '<a class="clipbd"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24">')
            WriteLine(fp, '<path path="#6D6D6D" fill="#6D6D6D" d="M19 2c-1.229 0-2.18-1.084-3-2h-8c-.82.916-1.771 2-3 2h-3v22h20v-22h-3zm-7 0c.552 0 1 .448 1 1s-.448 1-1 1-1-.448-1-1 .448-1 1-1zm8 20h-3.824c1.377-1.103 2.751-2.51')
            WriteLine(fp, '3.824-3.865v3.865zm0-8.457c0 4.107-6 2.457-6 2.457s1.518 6-2.638 6h-7.362v-18h4l2.102 2h3.898l2-2h4v9.543z"/></svg></a>')
            WriteLine(fp, OP.Name+'</h5>')
            if OP.Table != null then do
                self.WriteTable(OP.Table)
            end else if OP.Chart != null then do
                self.DrawChart(OP.Chart)
            end
            if OP.Footnote != null then 
                WriteLine(fp, '<div class="footnote">'+OP.Footnote+"</div>")
            WriteLine(fp, "</div>")
            
        end //loop over tables
        
        //End Sections
        if s2_name != null then WriteLine(fp, '</div>')
        if s1_name != null then WriteLine(fp, '</div>')
        if !TOA.NoHeader then WriteLine(fp, "</div>") //end main section 
    
    enditem //EndMethod - WriteTables
        
    //Write an individual table
    Macro "WriteTable" (TableOpts) do
    //TableOpts Values
    // .RowNames = Table Row Names (defaults to FT names)
    // .ColNames = Table column names (defaults to AT names) 
    //             - length must be cols or cols+1 to also include column name for row labels
    // .TableData = rows x columns array of string or numeric data to write to a table
    // .Formats = string to format all table data, or array of format strings matching TableData dimensions
    // .Class = Table HTML class value (defaults to "dataframe")
    // .TableStyle = stlye string to apply to table
    // .CellStyles = Array of style strings to apply to each cell
    
        shared UT
        TO = CopyArray(TableOpts) //Use a copy of the input opts
        fp = self.fp //quick reference
        
        //!!! TODO - better error checking of malformed input !!!
        
        // *** Some option processing ***
        //ColNames - set default, add first item as blank if needed
        if TO.ColNames = null then TO.ColNames = UT.Keys(self.Info.AT)+{"Total"}
        if TO.ColNames.length = TO.TableData[1].length then TO.ColNames = {''}+TO.ColNames
        
        //CellStyles: add first column (labels) to blank if needed
        if TypeOf(TO.CellStyles) = 'array' then do
            CellStyles = CopyArray(TO.CellStyles)
            if CellStyles[1].length = TO.TableData[1].length then do
                for ii = 1 to CellStyles.length do
                    CellStyles[ii] = {} + CellStyles[ii]
                end
            end
        end 
        else CellStyles = null
        
        
        //RowNames - set default
        if TO.RowNames = null then TO.RowNames = UT.Keys(self.Info.FT)+{"Total"}
        
        //Process number formats
        WriteData = CopyArray(TO.TableData)
        for ii = 1 to WriteData.length do
            for jj = 1 to WriteData[ii].length do
                //Load format string or array.  Use default for null values
                fmt = if TO.Formats = null then self.Formats.NumberFormats
                    else if TypeOf(TO.Formats) = 'array' then TO.Formats[ii][jj] 
                    else TO.Formats
                if fmt = null then fmt = self.Formats.NumberFormats
                
                //Perform conversion
                if TypeOf(WriteData[ii][jj]) = 'array' then 
                    WriteData[ii][jj] = self.ArrayString(WriteData[ii][jj])
                else if TypeOf(WriteData[ii][jj]) != 'string' then
                    WriteData[ii][jj] = Format(WriteData[ii][jj], fmt)
            end
            //Add row labels
            if TO.RowNames.length >= ii then do
                //InsertArrayElements(WriteData[ii], , {TO.RowNames[ii]})
                WriteData[ii] = {TO.RowNames[ii]} + WriteData[ii]
            end
        end
        
        //write the table
        cls = if TO.Class = null then ' class="dataframe"' else ' class="'+TO.Class+'"'
        sty = if TO.TableStyle = null then null else ' style="'+TO.TableStyle+'"'
        WriteLine(fp, '<table'+ cls + sty + '>')
        
        //Header row
        WriteLine(fp, '  <tr>')
        for t in TO.ColNames do
            WriteLine(fp, '    <th>'+t+'</th>')
        end
        WriteLine(fp, '  </tr>')
        
        //Data (includes row names)
        for ii = 1 to WriteData.length do
            WriteLine(fp, '  <tr>')
            for jj = 1 to WriteData[ii].length do
                td_sty = if CellStyles = null then null else ' style="'+CellStyles[ii][jj]+'"'
                WriteLine(fp, '    <td'+td_sty+'>'+WriteData[ii][jj]+'</td>')
            end
            WriteLine(fp, '  </tr>')
        end
        
        //end the table
        WriteLine(fp, '</table>')
    
    enditem //EndMethod - WriteTable
        
    //Write a single line as a title/header.
    Macro "WriteTitle" (Text, HLevel) do
    //Text = text to write
    //HLevel = header level to write, or 'b' for bold text (no header)
        
        if TypeOf(Text) != "string" then Return()
    
        if Lower(self.Formats.Filetype) = "html" then do
            fp = self.fp
            if TypeOf(HLevel) = 'string' and Lower(HLevel) = 'b' then 
                h = 'b'
            if TypeOf(HLevel) = "int" or TypeOf(HLevel) = "double" then
                h = String(HLevel)
            if TypeOf(h) != "string" then 
                h = '1'
                
            if h = 'b' then do
                sh = '<b>'
                eh = '</b><br>'
            end
            else do
                sh = '<h'+h+'>'
                eh = '</h'+h+'>'
            end
            
            WriteLine(fp, sh+Text+eh)
        
        end
        else do
            msg = "Unsupported file format requested"
            if TypeOf(self.Formats.Filetype) = "string" then
                msg = msg + " - " + self.Formats.Filetype
            ShowMessage(msg)
        end
    
    EndItem
    //EndMethod
    
    //Draw a chart using Chart.js - requires Chart.js library linked to html file
    Macro "DrawChart" (ChartOpts) do
    //ChartOpts 
    // .CanvasID = ID for canvas (must be kpet unique)
    // .Type = Type of chart (defaults to self.chartDefaults.Type)
    // .Labels = Array of data labels {'v1', 'v2', 'v3'} (ignored for scatter)
    // .Data = Array of data arrays: {{1, 2, 3}, {4, 5, 6}} (must be paired if scatter: {{{1, 2, 3}, {4, 5, 6}}, {{...}, {...}}})
    // .Colors = Array of bar chart colors, each an {rgb} array (defaults to self.chartOpts.ColorDefaults)
    // .Names = Array of dataset names (one for each dataset)  (one for each pair if scatter)
    // .Width = string Chart width (defaults to self.chartDefaults.Width)
    // .Height = string Chart width (self.chartDefaults.Height)
    //
    // -- Optional --
    // .XAxis = string with the X axis name
    // .YAxis = string with the Y axis name
    // .XMax = maximum value to plot
    // .YMax = maximum value to plot
    
        CO = CopyArray(ChartOpts) //Use a copy of the input opts
        fp = self.fp //quick reference
        
        //Defaults
        if CO.Width = null then CO.Width = self.chartDefaults.Width
        if TypeOf(CO.Width) != 'string' then CO.Width = String(CO.Width)
        if CO.Height = null then CO.Height = self.chartDefaults.Height
        if TypeOf(CO.Height) != 'string' then CO.Height = String(CO.Height)
        if CO.Colors = null then CO.Colors = self.chartDefaults.Colors
        if CO.Type = null then CO.Type = self.chartDefaults.Type
        
        WriteLine(fp, '<canvas id="'+CO.CanvasID+'" width="'+CO.Width+'" height="'+CO.Height+'" data-count="'+String(self.ChartCount)+'"></canvas>')
        WriteLine(fp, '<script>')
        WriteLine(fp, 'function dChart'+String(self.ChartCount)+' (canvasID) {')
        WriteLine(fp, 'var ctx = document.getElementById(canvasID);')
        WriteLine(fp, 'var data = {')
        if CO.Type = 'scatter' then do
            js_type = 'line'
            WriteLine(fp, '  datasets: [{')
            for ii = 1 to CO.Data.length do
                WriteLine(fp, '    label: "'+CO.Names[ii]+'",')
                WriteLine(fp, '    backgroundColor: "rgba('+JoinStrings(CO.Colors[ii]+{.4}, ", ")+')",')
                WriteLine(fp, '    borderColor: "rgba('+JoinStrings(CO.Colors[ii]+{1}, ", ")+')",')
                WriteLine(fp, '    borderWidth: 1,')
                WriteLine(fp, '    data: [')
                darr = null
                for jj = 1 to CO.Data[ii][1].length do
                    darr = darr + {'x: '+String(CO.Data[ii][1][jj]) + ', y: '+String(nz(CO.Data[ii][2][jj]))}
                end
                WriteLine(fp, "    {" + JoinStrings(darr, "}, {") + "}")
                
                 //end data and datasets array
                if ii < CO.Data.length then WriteLine(fp, '  ]}, {')
                else WriteLine(fp, "    ]}")
            end //datasets  
        end else do  //non-scatter
            if CO.Type = 'stacked' then js_type = 'bar'
            else js_type = CO.Type
            WriteLine(fp, '  labels: '+self.toJSarr(CO.Labels)+',')
            WriteLine(fp, '  datasets: [')
        
            for ii = 1 to CO.Data.length do
                WriteLine(fp, '  {')
                WriteLine(fp, '    label: "'+CO.Names[ii]+'",')
                WriteLine(fp, '    backgroundColor: "rgba('+JoinStrings(CO.Colors[ii]+{.4}, ", ")+')",')
                WriteLine(fp, '    borderColor: "rgba('+JoinStrings(CO.Colors[ii]+{1}, ", ")+')",')
                WriteLine(fp, '    borderWidth: 1,')
                WriteLine(fp, '    data: '+self.toJSarr(CO.Data[ii]))
                if ii < CO.Data.length then WriteLine(fp, '  },')
                else WriteLine(fp, '  }')
            end
        end
        WriteLine(fp, ']};') //end datasets, data, and line
        
        //Chart Options
        WriteLine(fp, 'var options = {scales: {yAxes: [{}], xAxes: [{}]}};')
        
        //Axis labels
        if CO.XAxis != null then do
            WriteLine(fp, 'options.scales.xAxes[0]["scaleLabel"] = {display: true, labelString: "'+CO.XAxis+'"};')
        end
        if CO.YAxis != null then do
            WriteLine(fp, 'options.scales.yAxes[0]["scaleLabel"] = {display: true, labelString: "'+CO.YAxis+'"};')
        end
        
        //Axis max/min
        if CO.XMax != null then
            WriteLine(fp, 'options.scales.xAxes[0]["ticks"] = {max: '+String(CO.XMax)+'};')
        if CO.YMax != null then
            WriteLine(fp, 'options.scales.yAxes[0]["ticks"] = {max: '+String(CO.YMax)+'};')
        
        //Extended types
        if CO.Type = 'scatter' then do
            WriteLine(fp, 'options.scales.xAxes[0]["type"] = "linear";')
            WriteLine(fp, 'options.scales.xAxes[0]["position"] = "bottom";')
            WriteLine(fp, 'options.showLines = false;')
        end else if CO.Type = 'stacked' then do
            WriteLine(fp, 'options.scales.yAxes[0]["stacked"] = true;')
        end
        
        WriteLine(fp, )
        WriteLine(fp, 'var chart = new Chart (ctx, {type: "'+js_type+'", data: data, options: options});')
        WriteLine(fp, '}')
        WriteLine(fp, 'dChart'+String(self.ChartCount)+'("'+CO.CanvasID+'");')
        WriteLine(fp, '</script>')
        
        self.ChartCount = self.ChartCount + 1
        
    
    enditem //EndMethod - DrawChart
    
    //Turn an array into a JavaScript string for use with DrawChart()
    // No nested arrays!
    Macro "toJSarr" (arr) do
    
        dim newArr[arr.length]
        for ii = 1 to arr.length do
            el = arr[ii]
            if TypeOf(el) = 'string' then 
                el = '"'+el+'"'
            else if TypeOf(el) != 'int' and TypeOf(el) != 'double' then
                Throw("toJSarr only works with strings and numbers.")
            
            newArr[ii] = el
        end
        
        rv = '['+JoinStrings(newArr, ", ")+']'
        Return(rv)
    
    enditem //EndMethod
    
    //Write a single blank line
    Macro "WriteBlank" do
        if Lower(self.Formats.Filetype) = "html" then do
            WriteLine(self.fp, '<br>')
        end
        else do
            msg = "Unsupported file format requested"
            if TypeOf(self.Formats.Filetype) = "string" then
                msg = msg + " - " + self.Formats.Filetype
            ShowMessage(msg)
        end
    EndItem
    //EndMethod
    
    //Convert an array to a formatted string, using recursion to show embedded arrays
	Macro "ArrayString" (InArr) do
		Arr = CopyArray(InArr) //Don't risk modifying the input array
		tmp = "{"
		for i = 1 to Arr.length do
			if TypeOf(Arr[i]) = "string" then do
				tmp = tmp + Arr[i] + ", "
			end
			else if TypeOf(Arr[i]) = "array" then do
				//tmp = tmp + "[subarray], "
                tmp = tmp + self.ArrayString(Arr[i]) + ", "
			end
			else if TypeOf(Arr[i]) = "int" then do
				tmp = tmp + format(Arr[i], "*.*") + ", "
			end
			else do
				tmp = tmp + format(Arr[i], "*.0*") + ", "
			end
		end
		tmp = left(tmp, len(tmp)-2) + "}" //Eliminate trailing comma, add }
		Return(tmp)
	
	enditem
	//EndMethod
    
    //Write a value to a cell by row and column number, format the cell
    Macro "WriteCell" (row, col, val, InOpts) do
    // Opts:
    //   string .Style = name of style (must be defined in Formats.excel)
    //   array .StyleOR = options array of style format overrides / additions
    
        //Process Input Opts
        if InOpts.Style != null and TypeOf(InOpts.Style) = "string" then
            style = self.Formats.excel.Styles.(InOpts.Style)
        else
            style = null

        if InOpts.StyleOR != null and TypeOf(InOpts.StyleOR) = "array" then
            style_or = InOpts.StyleOR
        else
            style_or = null
            
        //process the format
        frmt = InOpts.Format
        colwid = InOpts.ColWidth
    
        //Write value to cell
        CELL = self.XL.ActiveSheet.Cells[row][col]
        CELL.Value = val
        CELL.NumberFormat = self.ExcelFormat(frmt)
        CELL.ColumnWidth = colwid * 0.14

        //Apply style
        if style != null then
            self.SetProperty(CELL, style)
        if style_or != null then 
            self.SetProperty(CELL, style_or)
    
	enditem
	//EndMethod
    
    //Write an array of values to a row in Excel, format all cells in the row
    Macro "WriteLineXL" (row, col, val, InOpts) do
    // row: Row to write in Excel
    // col: Column to begin writing
    // val: one-dimensional array of values to write
    // Opts:
    //   string .Style = name of style (must be defined in Formats.excel)
    //   array .StyleOR = options array of style format overrides / additions
    
        //Process Input Opts
        if InOpts.Style != null and TypeOf(InOpts.Style) = "string" then
            style = self.Formats.excel.Styles.(InOpts.Style)
        else
            style = null

        if InOpts.StyleOR != null and TypeOf(InOpts.StyleOR) = "array" then
            style_or = InOpts.StyleOR
        else
            style_or = null
            
        //Write values to the row
        SHEET = self.XL.ActiveSheet
        RANGE = SHEET.Range(SHEET.Cells[row][col], 
                SHEET.Cells[row][col-1 + val.length])
        RANGE.Value = val
        
        //Apply style
        if style != null then
            self.SetProperty(RANGE, style)
        if style_or != null then 
            self.SetProperty(RANGE, style_or)
            
        //Apply number format (skip first column)
        if InOpts.Format != null then do
        RANGE = SHEET.Range(SHEET.Cells[row][col+1], 
                SHEET.Cells[row][col-1 + val.length])
            RANGE.NumberFormat = self.ExcelFormat(InOpts.Format)
            

        end //Format
        
        //release XL objects
        SHEET = null
        RANGE = null
    
	enditem
	//EndMethod
    
    //Recursive function to allow setting of object properties using
    //   an Opts array with multiple levels
    //
    Macro "SetProperty" (Obj, InOpts) do
    
        if InOpts != null and TypeOf(InOpts) = 'array' then do
            for _opt = 1 to InOpts.Length do
                key = InOpts[_opt][1]
                val = InOpts[_opt][2]
                
                if val = null or TypeOf(val) != 'array' then do
                
                    Obj.(key) = val
                end //Not an array
                else self.SetProperty(Obj.(key), val)

            end
        end
    
    enditem
    //EndMethod
    
	//**************************************************************************
	//** Run TLFD for a list of purposes/segments and write to a .bin file
	Macro "CalcTLFD" (trip_file, skim_file, skim_core, out_file, InOpts) do
	//** string trip_file = matrix file with trips.  TLFD will be reported for all cores.
	//** string skim_file = matrix file with skims.
	//** string skim_core = core name to use for impedance.
	//** string out_file = file where results should be saved
	//** Opts.MinBin = minimum bin number (defaults to 0)
	//** Opts.MaxBin = maximum bin number (defaults to 100)
	//** Opts.BinSize = bin size (defaults to 1)
	//** Opts.Tables = Array of core names to summarize (defaults to all cores)
	//** Opts.Period = String indicating period (for label only, ok to omit)
	//** Opts.Skims.<Core> = {File, core} //Skim file/core override by trip table core
	// self
	// Params/Opts are not type checked or verified to have valid contents.
	    
        //Process Opts
        MinBin = if InOpts.MinBin != null then InOpts.MinBin else 0
        MaxBin = if InOpts.MaxBin != null then InOpts.MaxBin else 100
        BinSize = if InOpts.BinSize != null then InOpts.BinSize else 1
        Tables = if InOpts.Tables != null then CopyArray(InOpts.Tables) else null
        per = if InOpts.Period != null then InOpts.Period else ""
        SkimOR = if InOpts.Skims != null then InOpts.Skims else null
        
        //Get core names
        mat = OpenMatrix(trip_file, )
        allcores = GetMatrixCoreNames(mat)
        mat = null
        
        //Filter cores to defined tables only
        if Tables != null then do
            cores = null
            for ii = 1 to allcores.length do
                if ArrayPosition(Tables, {allcores[ii]}, ) > 0 then cores = cores + {allcores[ii]}
            end
        end else cores = allcores
        
        //Get working directory
        t = SplitPath(out_file)
        dir = t[1] + t[2]
        
        //Initialize output table info
        Fields = {{"BIN", "Integer", 10, }}
        Vs = null
        Vs.Bin = Vector(MaxBin, "Long", {{"Sequence", 1, 1}})
        
        for _core = 1 to cores.length do
            core = cores[_core]
            
            //Set up table field name
            Fields = Fields + {{per+core, "Real", 10, 2}}
            
            //Check for skim override
            if SkimOR.(core) = null then do
                skim_usefile = skim_file
                skim_usecore = skim_core
            end else do
                t = SkimOR.(core)
                skim_usefile = t[1]
                skim_usecore = t[2]
            end
            
            //run TLD procedure
            Opts = null
            Opts.Input.[Base Currency] = {trip_file, core, , }
            Opts.Input.[Impedance Currency] = {skim_usefile, skim_usecore, , }
            Opts.Global.[Start Value] = MinBin
            Opts.Global.[End Value] = MaxBin
            Opts.Global.Size = BinSize
            Opts.Global.[Min Value] = 1 //ignore below min
            Opts.Global.[Max Value] = 0 //don't ignore over max
            Opts.Global.[Create Chart] = 0
            Opts.Output.[Output Matrix].Label = per + core+" TLFD"
            Opts.Output.[Output Matrix].Compression = 1
            tmp_file = dir + "__TEMP__TLFD_" + per + core + ".mtx"
            Opts.Output.[Output Matrix].[File Name] = tmp_file

            ret_value = RunMacro("TCB Run Procedure", "TLD", Opts, &Ret)
            if !ret_value then goto quit
            Ret = null
            
            //Load results into a vector, then delete temp file
            mat = OpenMatrix(tmp_file, )
            cur = CreateMatrixCurrency(mat, "TLD", , , )
            Vs.(per+core) = GetMatrixVector(cur, {{"Column", 1}})
            mat = null
            cur = null
            DeleteFile(tmp_file)
            
            
        end
        
        //Write results to a table
        tlfd_vw = CreateTable("TLFD", out_file, "FFB", Fields)
        AddRecords(tlfd_vw, , , {{"Empty Records", MaxBin}})
        SetDataVectors(tlfd_vw+"|", Vs, )
        CloseView(tlfd_vw)
        
        quit:
        
        Return(ret_value)
	
	
	enditem //EndMethod - CalcTLFD
    
    //Convert a hexidecimal color for use in Excel.  The Excel color property
    //  uses a decimal number value in reverse order. This macro takes a standard
    //  hex color value, reverses it, and converts it to decimal format
    //
    //  Any invalid input will silently fail and return black (0).
    Macro "ExcelHEX" (cHex) do

        if cHex = null or TypeOf(cHex) != 'string' then Return(0)
        if Len(cHex) != 6 then Return(0)

        //HEX Lookup
        HLU = {{'0', 0}, {'1', 1}, {'2', 2}, {'3', 3}, {'4', 4}, {'5', 5}, {'6', 6}, 
               {'7', 7}, {'8', 8}, {'9', 9}, {'A', 10}, {'B', 11}, {'C', 12}, 
               {'D', 13}, {'E', 14}, {'F', 15}}

        //Excel seems to work in reverse, so switch the order of the string
        cHex = cHex[5] + cHex[6] + cHex[3] + cHex[4] + cHex[1] + cHex[2]
        
        //Convert the HEX number to a decimal number
        tot = 0
        for ii = 0 to (Len(cHex) - 1) do
            place = Len(cHex) - (ii) //work backwards, 
            digit = cHex[place]      //Getting each digit
            val = HLU.(digit)        //Look up the decimal equivalent
            if val = null then Return(0) //Fail with black on invalid character
            
            tot = tot + val * Pow(16, ii)  //Then multiply by (16^ii)
        end
        
        Return(tot)
        
    enditem
    //EndMethod
    
    Macro "ExcelRGB" (xlR, xlG, xlB) do
    
        Return(Min(Max(nz(xlR), 0), 255) + 
               Min(Max(nz(xlG), 0), 255)*256 + 
               Min(Max(nz(xlB), 0), 255)*256*256)
               
    enditem
     //EndMethod
    
    //Convert a TransCAD Format string to an Excel format string.  This only
    //  works for specific pre-defined format strings - this macro can be
    //  extended to allow additional format strings.
    Macro "ExcelFormat" (TCfmt) do
        
        Formats = null
        Formats.("*,.") = "#,##0"
        Formats.("*0,") = "#,##0"
        Formats.("*0,.") = "#,##0"
        Formats.("*0,.0") = "#,##0.0"
        Formats.("*,0.") = "#,##0"
        Formats.("*0.00") = "#,##0.00"
        Formats.("*,0.00") = "#,##0.00"
        Formats.("*0.0%") = "0.0%"
        Formats.("*0.00%") = "0.00%"
        Formats.("*%0.0") = "0.0%"
 
        if TCfmt<>null then EXfmt = Formats.(TCfmt)

        Return(EXfmt)
        
    enditem
    //EndMethod
    
    Macro "WidthXL" (width) do
    
        Return((width - 5 )/ 7)
        
    enditem
    //EndMethod
    
EndClass



// ****************************************************************************************************************
// &&& Create Title Page***********************************************************************************************

Macro "RPT Title Page" (Perf)
    
    shared UT

	if Perf.Formats.Filetype = 'html' then do

		// Write the scenario information
		fp = Perf.fp
		SetCursor("Hourglass")
		WriteLine(fp,'<h1>SBTAM Summary Report</h1>')
        WriteLine(fp,'<div class="indent_h1">')
		WriteLine(fp,'<div class="titleInfo"><span class="blueText">Scenario Name: </span>' + Perf.Args.Info.Name + '</div>')
		WriteLine(fp,'<div class="titleInfo"><span class="blueText">Model Directory: </span>' + Perf.Args.Info.ModelDir + '</div>')
		WriteLine(fp,'<div class="titleInfo"><span class="blueText">Report File: </span>' + Perf.File + '</div>')
		WriteLine(fp,'<div class="titleInfo"><span class="blueText">Report Created on: </span>' + UT.FormatDate() + '</div>')
		WriteLine(fp,'<div class="titleInfo"><span class="blueText">Scenario Description: </span>' + Perf.Args.Info.Description + '</div>')
		WriteLine(fp, '</div>')
        
		//Write the table of contents
		WriteLine(fp,'<h2 id="folder_TOC"><span class="folder">&#x25b8; </span>Table of Contents</h2>')
        WriteLine(fp,'<div id="data_TOC" class="indent_h2">')
		WriteLine(fp, '<p class="grey">')
		for i = 1 to Perf.Report.length do
			Key = Perf.Report[i][1]
			Val = Perf.Report[i][2]
			Set = Perf.Settings.Report.(Key)
			
			if Set then WriteLine(fp,"<a href = \"#"+Val.Anchor+"\">")
			WriteLine(fp, "  " + string(i - 1) + ". " + Val.Contents)
			if Set then WriteLine(fp,"</a><br>")
			else WriteLine(fp,"<br>")
		end
        WriteLine(fp, '</div>')
	end //html
	
	else if Perf.Formats.Filetype = 'excel' then do
	
    
        //Rename the active worksheet
        //(First sheet since Title Page *should* always be called first and 
        //  must be called)
        Perf.XL.ActiveSheet.Name = "TitlePage"
        
        Perf.WriteCell(2, 3, "Summary Report for the SEMCOG Travel Model", 
                       {{"Style", "h1"}})
        Perf.WriteCell(5, 3, "Scenario Name: " + Perf.Args.Info.Name, {{"Style", "bluetext"}})   
        Perf.WriteCell(6, 3, "Scenario Directory: " + Perf.Args.Info.[Output Directory], 
                       {{"Style", "bluetext"}}) 
        Perf.WriteCell(7, 3, "Report File: " + Perf.File, {{"Style", "bluetext"}})   
        Perf.WriteCell(8, 3, "Report Created on: " + UT.FormatDate(), {{"Style", "bluetext"}})                

        //Save changes to the workbook
        Perf.XL.ActiveWorkbook.Save()
    
    end
    else do
        Throw("Invalid report file format specified")
    end
	
    
    Return(1)
EndMacro  //End of Title Page Macro

// ****************************************************************************************************************
// &&& Input File Data Summary

Macro "RPT Files" (Perf)
    
	//!!! !!! This should instead be read from Args, but Args does not 
    //        yet contain complete step descriptions
	//Define stage name titles
    Titles = {"Initialization and network creation",
             "Trip Generation",
             "Trip Distribution",
             "Mode Models",
             "Traffic Assignment",
             "Post-processing"}
	
	Stages = {"INI", "TGN", "DST", "MOD", "ASN", "PST"}
    
    
    DataTypes = {"Output", "Param", "Table",  "DbTable"} 
    DataHeaders = {"Output", "Parameters", "Tables", "Database Tables" } //nice names 
    
    //Create Input Table
    inputs = Perf.Args.Input
   
   //write input files
    InTable = null
    InRowNames = null
    for _in = 1 to inputs.length do
        InRowNames = InRowNames + {inputs[_in][1]}
        val = inputs[_in][2].Value
        if TypeOf(val) = "array" then val = Perf.ArrayString(val)
        InTable = InTable + {{val + "<br>"+inputs[_in][2].Desc}}
    end
    
    TB = null
    TB.Section1 = "Input Files"
    TB.Name = null //no table name for inputs
    TB.Table.RowNames = InRowNames
    TB.Table.ColNames = {'Key', 'Value & Description'}
    TB.Table.TableData = InTable
    TB.Table.Class = 'zebra'
    Tables = {CopyArray(TB)}

    //write the outputs, params, tables and dbtables for allstages
    for _stage = 1 to Stages.length do
        stage = Stages[_stage]
        title = Titles[_stage]
        
        for _type = 1 to DataTypes.length do 
            type = DataTypes[_type]
            outputs = Perf.Args.(type).(stage) //ouput/params/table/dbtable for each stage
            if outputs != null then do 
                Table = null
                Rows = null
                for _out = 1 to outputs.length do
                    Rows = Rows + {outputs[_out][1]}
                    val = outputs[_out][2].Value
                    if TypeOf(val) = "array" then val = Perf.ArrayString(val)
                    else if TypeOf(val) != "string" then val = String(val)
                    Table = Table + {{val+'<br>'+outputs[_out][2].Desc}}
                end
                TB = null
                TB.Section1 = Titles[_stage]
                TB.Name = DataHeaders[_type]
                TB.Table.RowNames = CopyArray(Rows)
                TB.Table.ColNames = {'Key', 'Value & Description'}
                TB.Table.TableData = CopyArray(Table)
                TB.Table.Class = 'zebra'
                Tables = Tables + {CopyArray(TB)}
                
            end //output!=null
        end //var
    end //stage
    


    Perf.WriteTables(Tables, )
       
    Return(True)
    

    
EndMacro


// ****************************************************************************************************************
// &&& Input Network Summary
//
// Input Network Summary
Macro "RPT Network Summary" (Perf)

    shared UT

    //FT and AT information
    ft_no = UT.Values(Perf.Info.FT)  //Returns a list of numbers only
    at_no = UT.Values(Perf.Info.AT)

    //Summary area - areas = {name, query}
    areas = Perf.ActiveAreas("Network")  //Can be Network or Zones
    
    //Define files
    dbd_file = Perf.Args.[Highway DB]

    //Dimension arrays to hold data: TableXXX[area][ft(row)][at(col)]
    dim TablesCL[areas.length]
    dim TablesLM[areas.length]
    TableGroup = {TablesCL, TablesLM}
    NameGroup = {"Network Centerline Summary", "Network Lane-Mile Summary"}
    
    //Open dbd network
    RunMacro("TCB Add DB Layers", dbd_file,,)
    layers = RunMacro("TCB get DB line and node layers", dbd_file)
    node_lyr = layers[1]
    link_lyr = layers[2]
    
    //loop over summary areas.
    for _area = 1 to areas.length do
        area_name = areas[_area][1]
        area_qry = areas[_area][2]
        SetView(link_lyr)
		
		// Create formula field FT based on AB_Facility_Type
		CreateExpression(link_lyr, "FT", "if CC = 1 then 10 else (if AB_Facility_Type<100 then  floor(AB_Facility_Type/10) else 99)", ) 	// SCAG: Need to add the field "CC" in the highway network
		
        //Do not report disabled links
        setcount = SelectByQuery("SummaryArea", "Several", 
                                 "Select * Where FT > 0", )
        setcount = SelectByQuery("SummaryArea", "Subset", area_qry, )
        //Only summarize if links are selected
        if setcount > 0 then do

            //Load data from view
            {FTv, ATv, CLv, 
            ABLANESv, BALANESv} = GetDataVectors(link_lyr+"|SummaryArea", 
                                  {"FT", "AB_AreaType", "Length", 				//SCAG: use AB_AreaType as the area type
                                  "AB_AMLANES", "BA_AMLANEs"}, )				//SCAG: use [AB/BA]-AMLANES as the number of lanes
                                  
            //Math
            LANESv = nz(ABLANESv) + nz(BALANESv)
            LMv = CLv*LANESv
                                  
            //Compute cross-class with marginals
            TablesCL[_area] = Perf.CrossTab(FTv, ATv, CLv, True)
            TablesLM[_area] = Perf.CrossTab(FTv, ATv, LMv, True)
                                  
        end //end if summary flag = 1 and setcount > 0
    end //end loop over summary areas
    CloseView(link_lyr)
   
    //Write the tables
    Tables = null
    TableNames = null
    for i = 1 to areas.length do
        for j = 1 to TableGroup.length do
            TB = null
            TB.Section1 = areas[i][1]
            TB.Name = NameGroup[j]
            TB.Table.TableData = TableGroup[j][i]
            
            Tables = Tables + {CopyArray(TB)}
        
        end
    end
    
    Perf.WriteTables(Tables, )
    
    Return(True)
    
EndMacro


Macro "RPT LU Data" (Perf)

    //Update progres bar
    shared UT

    //AT information
    at_no = UT.Values(Perf.Info.AT)  //Returns a list of numbers only
    at_names = UT.Keys(Perf.Info.AT)
    
    //Define files
    mdb_file = Perf.Args.Input.Database.Value
	socio_file = Perf.Args.Output.TGN.Socio.Value
	zdata_tname = Perf.Args.dbTable.TGN.zdata.Value
	arate_tname = Perf.Args.dbTable.TGN.Arate.Value

	//Open database tables
    //Open tables and files
	dsn_name = RunMacro("CreateDSN", mdb_file)
	socio_vw = OpenTable("SEData", "FFB", {socio_file})
	zdata_vw = RunMacro("OpenDSN", dsn_name, zdata_tname, , )
	arate_vw = RunMacro("OpenDSN", dsn_name, arate_tname, "SORT", )
	join_vw = JoinViews("SEData+Zdata", socio_vw+".TAZ", zdata_vw+".TAZ", )
	
	//Get list of fields and units
	{socio_flds, socio_units} = GetDataVectors(arate_vw+"|", {"LU_TYPE", "LU_UNIT"}, )
	socio_desc =V2A(socio_flds + " (" + socio_units + ")") //TYPE (UNITS)
	socio_flds = V2A(socio_flds)
	socio_units = V2A(socio_units)
	
	//Create an array to hold the LU summary table
	dim LUTable[socio_flds.length, at_no.length + 2]  //Summarize by AT + 2 for TOTAL and SOI
	
	//Read LU vectors and AT (Read all together, then separate the last vector)
	LU = GetDataVectors(join_vw+"|", socio_flds+{"AT"}, )
	AT = LU[LU.Length]
	LU = Subarray(LU, 1, LU.Length - 1)
	
	for i = 1 to socio_flds.length do
		for j = 1 to at_no.length do
			V = (if (AT = at_no[j]) then LU[i] else 0)
			LUTable[i][j] = VectorStatistic(V, "Sum", )
			
			//Accumulate total
			LUTable[i][at_no.length+2] = nz(LUTable[i][at_no.length+2]) + nz(LUTable[i][j])
			//Accumulate SOI Subtotal
			if j < 6 then do //!!! hard-coded reference to AT numbers for SOI
				LUTable[i][at_no.length+1] = nz(LUTable[i][at_no.length+1]) + nz(LUTable[i][j])
			end
		end
	end
    
    TB = null
    TB.Name = "Land Use Totals"
    
    TB.Table.TableData = LUTable
    TB.Table.RowNames = socio_desc
    TB.Table.ColNames = at_names+{"SOI Subtotal", "Total"}
    TB.Table.Class = "dataframe no-last-row"

    Tables = {TB}
    Perf.WriteTables(Tables, )
    
    Return(1)
    
EndMacro

// ****************************************************************************************************************
// &&& Trip Generation
Macro "RPT Trip Generation" (Perf)

	//Set up utilities
	shared UT

    //define files
    //pa_files = {Perf.Args.Output.TGN.PAbal.Value, Perf.Args.Output.TGN.Unbal.Value} //bal, unbal
	pa_files = {Perf.Args.[Peak Balanced PA], Perf.Args.[OffPeak Balanced PA]} 		//peak balanced PA, off-peak balanced PA
    pa_id = {"ID", "ID"} //PK and OP
    pa_fld = {"P", "A"}
    mdb_file = Perf.Args.Input.Database.Value			// Access database containing model input parameters and data
    //socio_file = Perf.Args.Output.TGN.Socio.Value		// Bivariate socioeconomic data
	socio_file = Perf.Args.Info.ModelDir + "SED\\model_sed_subregion.bin"
    frate_tname = Perf.Args.DbTable.TGN.Frate.Value		// Trip rate factors by K-district
    
    //define params
    //purp_names = Perf.Args.Table.TGN.Purp.Value
    //income_seg = Perf.Args.Table.TGN.ISeg.Value
    //inc_grps = Perf.Args.Table.TGN.Igrp.Value
    //inc_names = {"LI", "MI", "HI"}
	purp_names = { "HBWD", "HBWS", "HBSC", "HBCU", "HBSH", "HBSR",
				   "HBO",  "WBO",  "OBO",  "HBSP"}
	income_seg = { 1, 1, 0, 0, 1, 1, 1, 0, 0, 1}	// Income Segmentation Settings (1=yes, 0=no)
	inc_grps   = {"1", "2", "3", "4", "5"}			// Income definitions
	inc_names  = {"Inc1", "Inc2", "Inc3", "Inc4", "Inc5"}
	
	//purp_inc_names = { "HBWD1", "HBWD2", "HBWD3", "HBWD4", "HBWD5",
	//				   "HBWS1", "HBWS2", "HBWS3", "HBWS4", "HBWS5", 
	//				   "HBSC",
	//				   "HBCU",
	//				   "HBSH1", "HBSH2", "HBSH3", "HBSH4", "HBSH5",
	//				   "HBSR1", "HBSR2", "HBSR3", "HBSR4", "HBSR5", 
	//				   "HBO1", "HBO2", "HBO3", "HBO4", "HBO5", 
	//				   "WBO",
	//				   "OBO", 
	//				   "HBSP1", "HBSP2", "HBSP3", "HBSP4", "HBSP5"}


    areas = Perf.ActiveAreas("Zones")
    
	//purp_names2 --> Add income segmentation to purpose names
	purp_names2 = null
	purp_names2a = null
	for _purp = 1 to purp_names.length do
        ////HBU will not be income segmented
        //if purp_names[_purp] = "HBU" then do
        //    for _uname =  1 to u_names.length do
        //        purp_names2 = purp_names2 + {purp_names[_purp] + u_names[_uname]}
        //        purp_names2a = purp_names2a + {purp_names[_purp] + u_names[_uname]}
        //    end
        //end
		//else if !income_seg[_purp] and purp_names[_purp] <> "HBU" then do
		if !income_seg[_purp] then do
			purp_names2 = purp_names2 + {purp_names[_purp]} 
			purp_names2a = purp_names2a + {purp_names[_purp]}   //nice name for income seg
		end
		else do
			for _inc = 1 to inc_grps.length do
				purp_names2 = purp_names2 + {purp_names[_purp] + inc_grps[_inc]}
				purp_names2a = purp_names2a + {purp_names[_purp] + inc_names[_inc]}  //nice name (1 --> Inc1)
			end
		end
	end	

    //Open tables and files
    socio_vw = OpenTable("SEData", "FFB", {socio_file})
    SetView(socio_vw)
	// Create formula field "County" based on "CNTY", used to define area
	CreateExpression(socio_vw, "County", "if CNTY = 'Imperial' then 1 else (if CNTY = 'Los Angeles' then 2 else (if CNTY = 'Orange' then 3 else (if CNTY = 'Riverside' then 4 else (if CNTY = 'San Bernardino' then 5 else (if CNTY = 'Ventura' then 6 else 99)))))", ) 	// SCAG
    
	//SCAG: Summarize by area, bal for PK and OP
    for _area = 1 to areas.length do
        
		//Dimension TableData
		dim Data[purp_names2.length, 6] //cols: (1)prod, (2)attr, (3)p/HH, (4)p/POP, (5)% P, (6)% A
		
		for _pkop = 1 to 2 do

			//Load _pkop/un_pkop data
            pa_vw = OpenTable("PA", "FFB", {pa_files[_pkop],})
            join_vw = JoinViews("Join", pa_vw+"."+pa_id[_pkop], socio_vw+".SubregionTAZ", )
            SetView(join_vw)
            
            //and apply selection set
            cnt = SelectByQuery("Local", "Several", areas[_area][2], )
            if cnt > 0 then do
                cell_fmt = null
                p_tot = 0
                a_tot = 0
                
                //SED Basics
                {hh, pop} = GetDataVectors(join_vw+"|Local", {"HH", "POP"}, )		//SCAG SED fields
                
                hh = VectorStatistic(hh, "Sum", )
                pop = VectorStatistic(pop, "Sum", )
                
                for _purp = 1 to purp_names2.length do
                    purp = purp_names2[_purp]
                    {P, A} = GetDataVectors(join_vw+"|Local", {purp+"_P", purp+"_A"}, )
                    Data[_purp][1] = nz(Data[_purp][1]) + VectorStatistic(P, "Sum", )
                    Data[_purp][2] = nz(Data[_purp][2]) + VectorStatistic(A, "Sum", )
                    Data[_purp][3] = Data[_purp][1] / hh
                    Data[_purp][4] = Data[_purp][1] / pop
                    
                    //accumulate totals
                    p_tot = p_tot + Data[_purp][1]
                    a_tot = a_tot + Data[_purp][2]
                
                end //_purp
                
                
				if _pkop = 2 then do 			// only calculate totals and save table for writing when _pkop = 2 (op), Data[_purp][1] (productions) and Data[_purp][2] (attractions) are accumulating (PK + OP)
				
					//Add totals row
					Data = Data + {{p_tot, a_tot, p_tot / hh, p_tot / pop, , }}
					
					//Get percentages of totals, 
					for _purp = 1 to Data.length do
						Data[_purp][5] = Data[_purp][1] / p_tot
						Data[_purp][6] = Data[_purp][2] / a_tot
					end
					
					//Create cell format overrides
					for _purp = 1 to Data.length do
						cell_fmt = cell_fmt + {{, , "*0.00", "*0.00", "*0.0%", "*0.0%"}}
					end
                
					//Save table for writing
					TB = null
					basic_name = "Balanced Trips"
					TB.Name = basic_name + ' <span class="grey">('+areas[_area][1]+')</span>'
					TB.Section1 = areas[_area][1]
					if _pkop > 1 then TB.Section2 = "Additional Details"
					TB.Table.TableData = CopyArray(Data)
					TB.Table.RowNames = purp_names2a + {"All Trips"}
					TB.Table.ColNames = {"Purpose", "Productions", "Attractions", "Productions/HH", "Productions/Pop", "% of Productions", "% of Attractions"}
					TB.Table.Formats = cell_fmt
					TB.Table.Class = "dataframe no-last-col"
					
					Tables = Tables + {CopyArray(TB)}
				
					//Add a chart
					txData = TransposeArray(Data)
					chP = Subarray(txData[1], 1, txData[1].length-1)
					chA = Subarray(txData[2], 1, txData[2].length-1)
					
					CH = null
					basic_name = "Balanced Trips"
					CH.Name = basic_name + ' <span class="grey">('+areas[_area][1]+')</span>'
					CH.Section1 = areas[_area][1]
					if _pkop > 1 then CH.Section2 = "Additional Details"
					CH.Chart.CanvasID = 'chart_'+Substitute(basic_name+"_"+areas[_area][1], " ", "_",)
					CH.Chart.Type = 'bar'
					//CH.Chart.Labels = purp_names
					CH.Chart.Labels = purp_names2a
					CH.Chart.Data = {chP, chA}
					CH.Chart.Names = {"Productions", "Attractions"}
					
					Tables = Tables + {CopyArray(CH)}
				end
            
                CloseView(join_vw)
                CloseView(pa_vw)
            end
        
        end //pk / op
    end //area
    
    CloseView(socio_vw)
    
	//Write tables to file
    Perf.WriteTables(Tables,)	

/*    
    // ***************** Trip Generation by K-District *****************
    // (balanced only)
    bal = 1
    pa_vw = OpenTable("PA", "FFB", {pa_files[bal]})
    
    //Total Ps and As
    tmp = null
    for purp in purp_names do
        Ptmp = Ptmp + {"nz("+purp+"_P)"}
        Atmp = Atmp + {"nz("+purp+"_A)"}
    end
    Pexp = JoinStrings(Ptmp, " + ")
    Aexp = JoinStrings(Atmp, " + ")
    CreateExpression(pa_vw, "TOT_P", Pexp, )
    CreateExpression(pa_vw, "TOT_A", Aexp, )
    
    //Aggregate by KDIST
    agg_flds = {{"TOT_P", "sum", }, {"TOT_A", "sum", }}
    agg_vw = AggregateTable("AggPA", pa_vw+"|", "MEM", null, pa_vw+".KDIST", agg_flds, {{"Missing as zero"}})
    
    //Join rate factor table for K-District Names
    dsn_name = UT.CreateDSN(mdb_file)
    frate_vw = RunMacro("OpenDSN", dsn_name, frate_tname, "KDIST", )
    join_vw = JoinViews("Join", agg_vw+".KDIST", frate_vw+".KDIST", )
    
    //Create table
    KD = GetDataVectors(join_vw+"|", {agg_vw+".KDIST", "KDIST_NAME"}, )
    Vs = GetDataVectors(join_vw+"|", {"TOT_P", "TOT_A"}, )
    Data = null
    for ii = 1 to Vs.length do //start at 2, skipping KDIST name
        Data = Data + {V2A(Vs[ii])}
    end
    Data = TransposeArray(Data)
    
    //Save table for writing
    Rows = V2A(String(KD[1]) + ": " + KD[2])
    TB = null
    TB.Name = (if bal = 1 then "Balanced Trips" else "Unbalanced Trips") + " By K-District"
    TB.Section1 = "By K-District"
    TB.Table.TableData = CopyArray(Data)
    TB.Table.RowNames = Rows
    TB.Table.ColNames = {"K-District", "Productions", "Attractions"}
    TB.Table.Formats = "*,."
    TB.Table.Class = "dataframe no-last-col"
    
    Tables = Tables + {CopyArray(TB)}
    
    //Add a chart
    txData = TransposeArray(Data)
    
    CH = null
    basic_name = if bal = 1 then "Balanced Trips Chart" else "Unbalanced Trips Chart"
    CH.Name = basic_name + ' By K-District'
    CH.Section1 = "By K-District"
    CH.Chart.CanvasID = 'chart_'+Substitute(CH.Name, " ", "_",)
    CH.Chart.Type = 'bar'
    CH.Chart.Labels = V2A(String(KD[1]))
    CH.Chart.Data = txData
    CH.Chart.Names = {"Productions", "Attractions"}
    
    Tables = Tables + {CopyArray(CH)}
    

    
    
    //Write tables to file
    Perf.WriteTables(Tables,)
*/    
    Return(True)
EndMacro
//

// ****************************************************************************************************************
// &&& Trip Distribution
Macro "RPT Trip Distribution" (Perf)
    
	//skm_file = {Perf.Args.Output.DST.PKskm.Value, Perf.Args.Output.DST.OPskm.Value}
	skm_file = {Perf.Args.[Highway PK DA Skim], Perf.Args.[Highway OP DA Skim]}				//SCAG, used drive alone skim
	//pa_file  = {Perf.Args.Output.DST.PKdist.Value, Perf.Args.Output.DST.OPdist.Value}
	//PA person trip files
	pa_file = {{Perf.Args.[HBW PK Person Trips], 										//HBWD, HBWS by 5 income group
				Perf.Args.[HBNW PK Person Trips],										//HBSH, HBSR, HBSP and HBO by 5 income group
				Perf.Args.Info.ModelDir + "Tripdist\\Outputs\\HBSCPK_Trips.mtx",	//HBSC, HBSU 
				Perf.Args.[NHB PK Person Trips]},										//WBO, OBO
				
				{Perf.Args.[HBW OP Person Trips], 
				Perf.Args.[HBNW OP Person Trips],
				Perf.Args.Info.ModelDir + "Tripdist\\Outputs\\HBSCOP_Trips.mtx",
				Perf.Args.[NHB OP Person Trips]} }
				
    
	periods = {"Peak", "Off-Peak"}
	//skm_cores = {"[AB_PKTIME / BA_PKTIME] (Skim)", "[AB_OPTIME / BA_OPTIME] (Skim)"}
	skm_cores = {"NON-TOLL TIME", "NON-TOLL TIME"}										//SCAG, used non-toll time
	skmtt_cores = {"NON-TOLL TIME", "NON-TOLL TIME"}									//(with terminal time). SCAG might have only one time of travel time. Travel time with and without terminal time is set to be the same for now.
    skm_dist_cores = {"NON-TOLL DISTANCE", "NON-TOLL DISTANCE"}
	
    //Format strings for data columns
    fmts = {,,,"*0.0%", "*0.0", "*0.0", "*0.0", "*0.0", "*0.0"}
    Tables = null
    
    Dim Data[2] //for PK and OP, used to sum into daily at the end
    
    //use matrix cores to get purpose names
	purps = null
	//new_cores = null
	dim pa_file_cores[pa_file[1].length]
    for _pa_file = 1 to pa_file[1].length do 
		trp_mat = OpenMatrix(pa_file[1][_pa_file], )
		//new_cores = GetMatrixCoreNames(trp_mat)
		//purps = purps + new_cores
		purps = purps + GetMatrixCoreNames(trp_mat)
		//pa_file_cores[_pa_file] = new_cores.length
		trp_mat = null
		//new_cores = null
	end
	
	for _purp = 1 to purps.length do 	//remove the 'PK' at the end of a trip purpose, for example, HBWD1PK becomes HBWD1
		if Right(purps[_purp],2) = "PK" then purps[_purp] = Substring(purps[_purp],1,StringLength(purps[_purp])-2)
	end
	
	//dim purps[2]
	////purps = null
	//new_cores = null
	//dim pa_file_cores[2, pa_file[1].length]
	//for _per = 1 to periods.length do
	//	for _pa_file = 1 to pa_file[_per].length do 
	//		trp_mat = OpenMatrix(pa_file[_per][_pa_file], )
	//		new_cores = GetMatrixCoreNames(trp_mat)
	//		purps[_per] = purps[_per] + new_cores
	//		pa_file_cores[_per][_pa_file] = new_cores.length
	//		trp_mat = null
	//		new_cores = null
	//	end	
	//end	
	
	//purps = { "HBWD1", "HBWD2", "HBWD3", "HBWD4", "HBWD5",
	//	      "HBWS1", "HBWS2", "HBWS3", "HBWS4", "HBWS5", 
	//	      "HBSH1", "HBSH2", "HBSH3", "HBSH4", "HBSH5",
	//	      "HBSR1", "HBSR2", "HBSR3", "HBSR4", "HBSR5", 
	//	      "HBSP1", "HBSP2", "HBSP3", "HBSP4", "HBSP5",
	//	      "HBO1", "HBO2", "HBO3", "HBO4", "HBO5",
	//	      "HBSC",
	//	      "HBCU",
 	//	      "WBO",
	//	      "OBO"}
	
	//to accumulate daily intermediate values
	dim len_sumDY[purps.length]
	dim time_sumDY[purps.length]
	dim time_tt_sumDY[purps.length]	
    
    for _per = 1 to periods.length do
		
		Formats = null

        skm_mat = OpenMatrix(skm_file[_per], )
        skm_cur = CreateMatrixCurrency(skm_mat, skm_cores[_per], , , )
        skmtt_cur = CreateMatrixCurrency(skm_mat, skmtt_cores[_per], , , )
        skmlen_cur = CreateMatrixCurrency(skm_mat, skm_dist_cores[_per], , , )
		
		////use matrix cores to get purpose names
		//purps = null
		//new_cores = null
		//dim pa_file_cores[pa_file[1].length]
		//for _pa_file2 = 1 to pa_file[1].length do 
		//	trp_mat = OpenMatrix(pa_file[1][_pa_file], )
		//	new_cores = GetMatrixCoreNames(trp_mat)
		//	purps = purps + new_cores						//Chao: reduced the dimentions for purps and pa_file_cores
		//	pa_file_cores[_pa_file2] = new_cores.length
		//	trp_mat = null
		//	new_cores = null
		//end
		//
		//if _per = 1 then do //initialize
		//	//to accumulate daily intermediate values
		//	dim len_sumDY[purps.length]
		//	dim time_sumDY[purps.length]
		//	dim time_tt_sumDY[purps.length]		
		//end

		n_purp = 0			// used to track the position in purps
		
		dim last_line[9]	// to accumulate for each time period
		len_sum = 0
		time_sum = 0
		time_tt_sum = 0		
		
        for _pa_file = 1 to pa_file[1].length do			//for each pa matrix file
			
			trp_mat = OpenMatrix(pa_file[_per][_pa_file], )
	
			t = SplitPath(pa_file[_per][_pa_file])
			Opts = null
			Opts.[File Name] = t[1]+t[2]+"__TEMP__SkimRpt.mtx"
			Opts.Label = "Scratch"
			Opts.Tables = {"DATA"}
			Opts.[Memory Only] = "True"
			tmp_mat = CopyMatrixStructure({skm_cur}, Opts)
			tmp_cur = CreateMatrixCurrency(tmp_mat, "DATA", , , )
			
			purps2 = GetMatrixCoreNames(trp_mat)
			
			//for _purp = 1 to pa_file_cores[_pa_file] do 	//for each core in the matrix
			for _purp = 1 to purps2.length do 	//for each core in the matrix
				
				n_purp = n_purp + 1
				purp = purps2[_purp]
				dim line[9]
				//(1)trips (2)intra (3)inter (4)%inter (5)miles (6)min (7)mph (8) min(term) (9)mph(term)
				trp_cur = CreateMatrixCurrency(trp_mat, purp, , , )
			
				//Total and intrazonal trips
				line[1] = VectorStatistic(GetMatrixVector(trp_cur, {{"Marginal", "Row Sum"}}), "Sum", )
				line[2] = VectorStatistic(GetMatrixVector(trp_cur, {{"Diagonal", "Column"}}), "Sum", )
				line[3] = line[1] - line[2]
				line[4] = line[2] / line[1]
				
				//Time and speed
			
				//Avg Length (Miles)
				tmp_cur := skmlen_cur * trp_cur
				V = VectorStatistic(GetMatrixVector(tmp_cur, {{"Marginal", "Row Sum"}}), "Sum", )
				line[5] = V / line[1]
				len_sum = len_sum + V
				//len_sumDY[_purp] = nz(len_sumDY[_purp]) + V
				len_sumDY[n_purp] = nz(len_sumDY[n_purp]) + V
				
				//Avg time (no termainl time)
				tmp_cur := skm_cur * trp_cur
				V = VectorStatistic(GetMatrixVector(tmp_cur, {{"Marginal", "Row Sum"}}), "Sum", )
				line[6] = V / line[1]
				time_sum = time_sum + V
				//time_sumDY[_purp] = nz(time_sumDY[_purp]) + V
				time_sumDY[n_purp] = nz(time_sumDY[n_purp]) + V
				
				//Avg Speed (no term)
				line[7] = line[5] / line[6] * 60
				
				//Avg time (with termainl time)
				tmp_cur := skmtt_cur * trp_cur
				V = VectorStatistic(GetMatrixVector(tmp_cur, {{"Marginal", "Row Sum"}}), "Sum", )
				line[8] = V / line[1]
				time_tt_sum = time_tt_sum + V
				//time_tt_sumDY[_purp] = nz(time_tt_sumDY[_purp]) + V
				time_tt_sumDY[n_purp] = nz(time_tt_sumDY[n_purp]) + V
				
				//Avg Speed (no term)
				line[9] = line[5] / line[8] * 60
				
				
				//Add line to table, address formatting
				Data[_per] = Data[_per] + {CopyArray(line)}
				Formats = Formats + {fmts}
				
				//Sum totals
				last_line[1] = nz(last_line[1]) + line[1]
				last_line[2] = nz(last_line[2]) + line[2]
				last_line[3] = nz(last_line[3]) + line[3]
				last_line[4] = nz(last_line[2]) / last_line[1]
				
				//Close innermost loop matrix currency
				trp_cur = null
			end	//end of _purp
			
			trp_mat = null
			tmp_mat = null
			tmp_cur = null
			
		end	//end of _pa_file	
			
		//Calculate total time and speed
		last_line[5] = len_sum / last_line[1]
		last_line[6] = time_sum / last_line[1]
		last_line[7] = last_line[5] / last_line[6] * 60
		last_line[8] = time_tt_sum / last_line[1]
		last_line[9] = last_line[5] / last_line[8] * 60
		
		Data[_per] = Data[_per] + {CopyArray(last_line)}
		Formats = Formats + {fmts}
		
        TB = null
        TB.Section1 = periods[_per]
        TB.Name = "Trip Distribution Summary"
        TB.Footnote = "* Includes terminal time"
        TB.Table.TableData = CopyArray(Data[_per])
        TB.Table.Formats = CopyArray(Formats)
        TB.Table.RowNames = purps + {"All Trips"}
        TB.Table.ColNames = {"Purpose", "Trips", "Intrazonal", "Interzonal", "% Intrazonal", "Avg Length (miles)", "Avg Time (min)", "Avg Speed", "Avg Time (min)*", "Avg Speed*"}
        TB.Table.Class = "dataframe no-last-col"
        
        Tables = Tables + {CopyArray(TB)}
        
 
        skm_mat = null
        skm_cur = null
        skmtt_cur = null
        skmlen_cur = null

    
    end //_per
	
    //Add daily table by summing PK and OP
    len_sumDY = len_sumDY + {Sum(len_sumDY)}
    time_sumDY = time_sumDY + {Sum(time_sumDY)}
    time_tt_sumDY = time_tt_sumDY + {Sum(time_tt_sumDY)}
    DataDY = null
    for _purp = 1 to Data[1].length do
        dim line[9]
        for ii = 1 to 3 do
            line[ii] = Data[1][_purp][ii] + Data[2][_purp][ii]
        end
        line[4] = line[2] / line[1]
        
        //Speeds and times
        
        line[5] = len_sumDY[_purp] / line[1]
        line[6] = time_sumDY[_purp] / line[1]
        line[7] = line[5] / line[6] * 60
        line[8] = time_tt_sumDY[_purp] / line[1]
        line[9] = line[5] / line[8] * 60
        
        DataDY = DataDY + {CopyArray(line)}
        
    end
    
    TB = null
    TB.Section1 = "Daily"
    TB.Name = "Trip Distribution Summary"
    TB.Footnote = "* Includes terminal time"
    TB.Table.TableData = CopyArray(DataDY)
    TB.Table.Formats = CopyArray(Formats) //remains from pk / op
    TB.Table.RowNames = purps + {"All Trips"}
    TB.Table.ColNames = {"Purpose", "Trips", "Intrazonal", "Interzonal", "% Intrazonal", "Avg Length (miles)", "Avg Time (min)", "Avg Speed", "Avg Time (min)*", "Avg Speed*"}
    TB.Table.Class = "dataframe no-last-col"
    
    Tables = {CopyArray(TB)} + Tables

    Perf.WriteTables(Tables, )
    
    Return(True)
EndMacro

Macro "RPT Trip Length Frequencies" (Perf)
    
	//Params
    max_plot_time = 40 //Maximum time to include in plots
    //purp_names = Perf.Args.Table.TGN.Purp.Value
	//income_seg = Perf.Args.Table.TGN.ISeg.Value
	//inc_grps   = Perf.Args.Table.TGN.Igrp.Value
	
    Verbose = True //!!! Must use verbose until TLFD is updated to use memory matrix, or summing is done in Trip Dist instead

	//skm_files = {Perf.Args.Output.DST.PKskm.Value, Perf.Args.Output.DST.OPskm.Value}
	skm_files = {Perf.Args.[Highway PK DA Skim], Perf.Args.[Highway OP DA Skim]}				//SCAG, used drive alone skim
    t = SplitPath(skm_files[1])
    pth = t[1]+t[2]
	//pa_files  = {Perf.Args.Output.DST.PKdist.Value, Perf.Args.Output.DST.OPdist.Value}
	pa_files = {{Perf.Args.[HBW PK Person Trips], 										//HBWD, HBWS by 5 income group
				Perf.Args.[HBNW PK Person Trips],										//HBSH, HBSR, HBSP and HBO by 5 income group
				Perf.Args.Info.ModelDir + "Tripdist\\Outputs\\HBSCPK_Trips.mtx",	//HBSC, HBSU 
				Perf.Args.[NHB PK Person Trips]},										//WBO, OBO
				
				{Perf.Args.[HBW OP Person Trips], 
				Perf.Args.[HBNW OP Person Trips],
				Perf.Args.Info.ModelDir + "Tripdist\\Outputs\\HBSCOP_Trips.mtx",
				Perf.Args.[NHB OP Person Trips]} }
	
    tlfd_files = {pth+"PK_TLFD.bin", pth+"OP_TLFD.bin"}
	
	
    ////use matrix cores to get purpose names
	//purps = null
	////new_cores = null
	//dim pa_file_cores[pa_file[1].length]
    //for _pa_file = 1 to pa_file[1].length do 
	//	trp_mat = OpenMatrix(pa_file[1][_pa_file], )
	//	//new_cores = GetMatrixCoreNames(trp_mat)
	//	//purps = purps + new_cores
	//	purps = purps + GetMatrixCoreNames(trp_mat)
	//	//pa_file_cores[_pa_file] = new_cores.length
	//	trp_mat = null
	//	//new_cores = null
	//end
	//
	//for _purp = 1 to purps.length do 	//remove the 'PK' at the end of a trip purpose, for example, HBWD1PK becomes HBWD1
	//	if Right(purps[_purp],2) = "PK" then purps[_purp] = Substring(purps[_purp],1,StringLength(purps[_purp])-2)
	//end

	
	
	purp_names = { "HBWD", "HBWS", "HBSH", "HBSR", "HBSP", "HBO",  
				   "HBSC", "HBCU", "WBO",  "OBO" }
	
	num_array = {"0","1","2","3","4","5","6","7","8","9"}	//used to determine if a letter is a number
	
    
    //Matrix indices - row/col index (null to use default/only index)
    skm_idx = {,}
    trp_idx = {,} 
    
	periods = {"PK", "OP", "DY"}
    per_names = {"Peak", "Off-Peak", "Daily"}
	//skmlen_cores = {"Length (Skim)", "Length (Skim)"}
	//skmtt_cores = {"[AB_PKTIME / BA_PKTIME] (Skim with TT)", "[AB_OPTIME / BA_OPTIME] (Skim with TT)"}
	skmlen_cores = {"NON-TOLL DISTANCE", "NON-TOLL DISTANCE"}
	skmtt_cores = {"NON-TOLL TIME", "NON-TOLL TIME"}			//Travel time with terminal time.

    
    //Format strings for data columns
    Tables = null
    dim TRIPS_DY[purp_names.length]
    
    for _per = 1 to periods.length do
        per = periods[_per]

        if per != "DY" then do

            skm_mat = OpenMatrix(skm_files[_per], )
            skm_curs = CreateMatrixCurrencies(skm_mat, skm_idx[1], skm_idx[2], )
            
            for _pa_file = 1 to pa_files[1].length do			//for each pa matrix file
				
				trp_mat_seg = OpenMatrix(pa_files[_per][_pa_file], )
				trp_cur_seg = CreateMatrixCurrencies(trp_mat_seg, trp_idx[1], trp_idx[2], )
				
				if _pa_file = 1 then do 						//create a tempory file, with all trip purposes (not by income group) + skim length + skim travel time
					//Add up segments into a temporary matrix
					//!!! !!! TODO: Consider creating this summary as part of the 
					//              Trip Distribution Model instead
					t = SplitPath(pa_files[_per][_pa_file])
					Opts = null
					scratch_file = t[1]+t[2]+"__TEMP__"+per+"_Trips"+t[4]
					Opts.[File Name] = scratch_file
					Opts.Label = per + " PA Trips by Purpose"
					Opts.Tables = purp_names + {skmlen_cores[_per], skmtt_cores[_per]} //include skims and a scratch core in this mem matrix for improved speed
					if !Verbose then Opts.[Memory Only] = "True"
					else Opts.Compression = 1
					trp_mat = CopyMatrixStructure({trp_cur_seg[1][2]}, Opts)
					trp_curs = CreateMatrixCurrencies(trp_mat, , , )
				end
				
				purp_cores = GetMatrixCoreNames(trp_mat_seg)
				
				//Accumulate trips by purpose (aggregate income groups)
				//purp_current = null
				for _purp = 1 to purp_cores.length do
					purp2 = purp_cores[_purp]
					if Right(purp2,2) = "PK" or Right(purp2,2) = "OP" then do 
						purp = Substring(purp2,1,StringLength(purp2)-2)			//now "HBWD1PK" becomes "HBWD1"
						last_letter = Right(purp,1)
						if ArrayPosition(num_array,{last_letter},) > 0 then do 
							purp = Substring(purp,1,StringLength(purp)-1)		//now "HBWD1" becomes "HBWD"
						end
					end
					else purp = purp2
					
					trp_curs.(purp) := nz(trp_curs.(purp)) + nz(trp_cur_seg.(purp2))
					
					//if purp = purp_current then do 
					//	trp_curs.(purp) := nz(trp_curs.(purp)) + nz(trp_cur_seg.(purp2))
					//else do 
					//	purp_current = 	purp
					//	trp_curs.(purp) := nz(trp_curs.(purp)) + nz(trp_cur_seg.(purp2))
					//end	
					//
					//
					//purp_current = 	purp
					//
					//trp_curs.(purp) := nz(trp_curs.(purp)) + nz(trp_cur_seg.(purp2))	
                    //
                    //
					//
					//	if 
					//if income_seg[_purp] then do
					//	for inc in inc_grps do
					//		trp_curs.(purp) := nz(trp_curs.(purp)) + nz(trp_cur_seg.(inc+purp))
					//	end //inc
					//end //segmented
					//else do
					//	trp_curs.(purp) := trp_cur_seg.(purp)
					//end
				end //for
				trp_mat_seg = null
				trp_cur_seg = null
			end		//for each pa matrix file

			//Add skim data to MEM matrix
			trp_curs.(skmlen_cores[_per]) := skm_curs.(skmlen_cores[_per])
			trp_curs.(skmtt_cores[_per]) := skm_curs.(skmtt_cores[_per])
			
			//Done with segmented matrix and skim, so close
			skm_mat = null
			skm_curs = null
            
			//Calculate trip length frequency distributions
			Opts = null
			Opts.Tables = purp_names
			ok = Perf.CalcTLFD(scratch_file, scratch_file, skmtt_cores[_per], tlfd_files[_per], Opts)
			if !ok then Throw("Batch Error calculating TLFD.  See the log file for details.")
				
			//Close scratch matrix
			trp_mat = null
			trp_curs = null	
				
			//Create TLFD Charts and tables
			
			//Load from TLFD
			tlfd_vw = OpenTable("TLFD", "FFB", {tlfd_files[_per]})
			Vs = GetDataVectors(tlfd_vw+"|", {"BIN"}+purp_names, )
			BINS = Vs[1]
			//BINS_lim = SubVector(BINS, 1, max_plot_time)
			BINS_lim = a2v(Subarray(v2a(BINS), 1, max_plot_time))		//SCAG, SubVector is not available in TransCAD v6
			TRIPS = Subarray(Vs, 2, )
			
			//Accumulate daily trips by bin
			for ii = 1 to TRIPS.length do
				TRIPS_DY[ii] = nz(TRIPS_DY[ii]) + TRIPS[ii]
			end

        end //if not daily
        else do
            TRIPS = CopyArray(TRIPS_DY)
        end
        
        //Get percentages, limit to intended scope
        dim PCTS[TRIPS.length]
        dim PCTS_lim[TRIPS.Length]
        dim TRIPS_lim[TRIPS.Length]
        for ii = 1 to TRIPS.length do
            PCTS[ii] = nz(TRIPS[ii] / VectorStatistic(TRIPS[ii], "Sum", ))
            //TRIPS_lim[ii] = SubVector(TRIPS[ii], 1, max_plot_time)
			TRIPS_lim[ii] = a2v(Subarray(v2a(TRIPS[ii]), 1, max_plot_time))		//SCAG, SubVector is not available in TransCAD v6
            //PCTS_lim[ii] = SubVector(PCTS[ii], 1, max_plot_time)
			PCTS_lim[ii] = a2v(Subarray(v2a(PCTS[ii]), 1, max_plot_time))		//SCAG, SubVector is not available in TransCAD v6
        end
        
        
        //Chart
        CH = null
        CH.Name = "Trip Length Frequency Distribution ("+per+")"
        CH.Section1 = per_names[_per]
        CH.Footnote = "Click the legend to show/hide trip purposes"
        CH.Chart.CanvasID = "tlfd_chart_"+per
        CH.Chart.Type = "line"
        CH.Chart.Labels = String(BINS_lim - 1) + "-" + String(BINS_lim)
        CH.Chart.Data = PCTS_lim
        CH.Chart.Width = 800
        CH.Chart.Height = 400
        CH.Chart.Names = purp_names
        CH.Chart.XAxis = "Travel Time (min)"
        CH.Chart.YAxis = "Percent of Trips"
        
        
        //... and table
        TOT = 0 //all purposes
        for ii = 1 to TRIPS.length do
            TOT = TOT + nz(TRIPS[ii])
        end
        dim Data[TRIPS.length]
        for ii = 1 to Data.length do
            Data[ii] = V2A(TRIPS[ii])
        end
        Data = Data + {V2A(TOT)}
        Data = TransposeArray(Data)
        TB = null
        TB.Name = "Trip Length Frequency Distribution Table ("+per+")"
        TB.Section1 = per_names[_per]
        TB.Section2 = "Data Table"
        TB.Table.RowNames = V2A(String(BINS - 1) + "-" + String(BINS))
        TB.Table.ColNames = {"Time (min)"} + purp_names + {"Total"}
        TB.Table.TableData = Data


    if per = "DY" then
        Tables = {CH, TB} + Tables
    else
        Tables = Tables + {CH, TB}

    end //_per

    Perf.WriteTables(Tables, )
    
    Return(True)
EndMacro

// ****************************************************************************************************************
// &&& Mode Choice Summary
//  Summarizes results of mode choice
//
Macro "RPT Mode Choice" (Perf)

    //NOTE: This uses a different AREA query than most reports
    //!!! Consider revising !!!

    //Define files
	sum_file = Perf.Args.Output.MOD.ModeSummary.Value
				
    area_names = {"ALL", "SOI"}
    area_secs  = {"Entire Model", "SOI Only"}
    pers = {"PK", "OP", "DY"}
    per_names = {"Peak Period ", "Off-Peak", "Daily"} //DY/Daily must be last for sum
    
    sum_vw = OpenTable("Summary", "FFB", {sum_file}, )
    SetView(sum_vw)
    
    ColNames = GetFields(sum_vw, "All")
    ColNames = ColNames[1]
    ColNames = ExcludeArrayElements(ColNames, 2, 3) + {"Total"} //Exclude Income, Area, and Period field names
    
    Tables = null
    for _area = 1 to area_names.length do
        area = area_names[_area]
        dim recs[pers.length]
        AreaTables = null
        for _per = 1 to pers.length do
            per = pers[_per]
            
            if per = "DY" then do
                recs[_per] = CopyArray(recs[1])
                for _source = 2 to (pers.length - 1) do
                    for ii = 1 to recs[_per].length do //start at 1 (all rows)
                        for jj = 1 to recs[_per][ii].length do //start at 2 (seonc)
                            recs[_per][ii][jj] = nz(recs[_per][ii][jj]) + nz(recs[_source][ii][jj])
                        end //ii (cols)
                    end //jj (rows)
                end
            end else do
                cnt = SelectByQuery("S", "Several", 'Select * Where Area = "'+area+'"')
                cnt = SelectByQuery("S", "Subset", 'Select * Where Period = "'+per+'"')
                recs[_per] = GetRecordsValues(sum_vw+"|S", GetFirstRecord(sum_vw+"|S", ), , , cnt, "Row", )
                
                //Move first column into separate RowNames
                //and add a total column
                dim RowNames[recs[_per].length]
                for ii = 1 to recs[_per].length do
                    RowNames[ii] = recs[_per][ii][1]
                    if recs[_per][ii][2] != "ALL" then
                        RowNames[ii] = RowNames[ii] + recs[_per][ii][2]
                        
                    recs[_per][ii] = ExcludeArrayElements(recs[_per][ii], 1, 4) //Exclude purpose, income, period, area
                end
                RowNames = RowNames + {"Total"}
            end
            
            //Add totals
            recs_marg = Perf.Marginals(recs[_per])
            
            //Get percentages
            dim recs_pct[recs[_per].length]
            for ii = 1 to recs[_per].length do
                V = A2V(recs[_per][ii])
                V = Round(V / VectorStatistic(V, "Sum", ), 4)
                recs_pct[ii] = V2A(V)
            end
            
            //Percentage totals, special consideration for last row
            recs_pct_marg = Perf.Marginals(recs_pct)
            ii = recs_pct_marg.length
            for jj = 1 to recs_pct_marg[ii].length do
                recs_pct_marg[ii][jj] = recs_marg[ii][jj] / recs_marg[ii][recs_marg[ii].length]
            end
                        
            //Table
            TB = null
            TB.Name = per_names[_per] + " Mode Chocie"+ ' <span class="grey">('+area_secs[_area]+')</span>'
            TB.Section1 = area_secs[_area]
            TB.Table.TableData = recs_marg
            TB.Table.RowNames = RowNames
            TB.Table.ColNames = ColNames
            
            //Percent Table
            TBP = null
            TBP.Name = per_names[_per] + " Mode Chocie Share"+ ' <span class="grey">('+area_secs[_area]+')</span>'
            TBP.Section1 = area_secs[_area]
            TBP.Table.TableData = recs_pct_marg
            TBP.Table.RowNames = RowNames
            TBP.Table.ColNames = ColNames
            TBP.Table.Formats = "*.0%"
            
            //...and chart
            ChartData = TransposeArray(recs_pct)  //use the % shares, no marginals

            CH = null
            CH.Name = per_names[_per] + " Mode Chocie Chart"+ ' <span class="grey">('+area_secs[_area]+')</span>'
            CH.Section1 = area_secs[_area]
            CH.Chart.CanvasID = "mode_choice_"+per_names[_per]+"_"+Substitute(area_secs[_area], " ", "_", )
            CH.Chart.Type = 'stacked'
            CH.Chart.Labels = Subarray(RowNames, 1, RowNames.length-1) //remove total
            CH.Chart.Data = ChartData
            CH.Chart.Names = Subarray(ColNames, 2, ColNames.length-2)  //remove purp name and total
            CH.Chart.Width = 800
            CH.Chart.YAxis = "Share of Trips"
            CH.Chart.YMax = 1
            
            if per = "DY" then do //Put the DY table first
                AreaTables = {CopyArray(TB), CopyArray(TBP), CopyArray(CH)} + AreaTables
            end else do //Collapse period tables
                TB.Section2 = "By Period"
                CH.Section2 = "By Period"
                AreaTables = AreaTables + {CopyArray(TB), CopyArray(TBP), CopyArray(CH)}
            end
        end //_per
        
        Tables = Tables + AreaTables
        
    end //area
		

    Perf.WriteTables(Tables, )
    Return(True)
	
EndMacro

// ****************************************************************************************************************
// &&& Assigned Vehicle Trips
//
Macro "RPT Assigned Vehicle Trips" (Perf)

    areas = Perf.ActiveAreas("Zones")
    
    //Define files
	periods = {"AM Peak", "PM Peak", "Off Peak"}
	od_files = {Perf.Args.Output.ASN.AMtrips.Value, Perf.Args.Output.ASN.PMtrips.Value, Perf.Args.Output.ASN.OPtrips.Value}
	pabal_file = Perf.Args.Output.TGN.PAbal.Value //Balanced PA file to read AT / Summary area info
	
    //use matrix cores to get purpose names
    mat = OpenMatrix(od_files[1], )
    purps = GetMatrixCoreNames(mat)
    purps = Subarray(purps, 1, purps.length-1) //exclude final total core
    mat = null
	
    //Use pabal to select by summary area
	pabal_vw = OpenTable("PABalanced", "FFB", {pabal_file,})
    SetView(pabal_vw)
    
        
    //5 rows: (1)Intra, (2)iter orig, (3)tot orig, (4)inter dest, (5)tot dest
    //purps.length columns (will add total)
    dim TableData[areas.length, periods.length, 5, purps.length]
    
    //Summarize by O and D in each area
    for _per = 1 to periods.length do
        per = periods[_per]
        //Open the matrix, copy to memory
        tmat = OpenMatrix(od_files[_per], )
        tcur = CreateMatrixCurrency(tmat, purps[1], , , )
        t = SplitPath(od_files[_per])
        Opts = null
        Opts.[File Name] = t[1]+t[2]+"__TEMP__"+t[3]+t[4]
        Opts.Label = per + "OD MEM Matrix"
        Opts.[Memory Only] = "True"
        mat = CopyMatrix(tcur, Opts)
        all_idx = GetMatrixBaseIndex(mat)
        tmat = null
        tcur = null
        
        for _area = 1 to areas.length do 

        	SelectByQuery("Local", "Several", areas[_area][2],)
        	idx = CreateMatrixIndex(areas[_area][1], mat, "Both", pabal_vw+"|Local", "TAZ", )
            
            for _purp = 1 to purps.length do
                purp = purps[_purp]
                //Intrazonal trips
				cur = CreateMatrixCurrency(mat, purp, idx, idx, )
                TableData[_area][_per][1][_purp] = VectorStatistic(GetMatrixVector(cur, {{"Diagonal", "Row"}}), "Sum", )
				cur = null
                
                //Total Trips: Origin in juris
                cur = CreateMatrixCurrency(mat, purp, idx, all_idx[2], )
                TableData[_area][_per][3][_purp] = VectorStatistic(GetMatrixVector(cur, {{"Marginal", "Row Sum"}}), "Sum", )
                TableData[_area][_per][2][_purp] = TableData[_area][_per][3][_purp] - TableData[_area][_per][1][_purp]  //tot - intra
                cur = null
                
                //Total Trips: Destination in juris
                cur = CreateMatrixCurrency(mat, purp, all_idx[1], idx, )
                TableData[_area][_per][5][_purp] = VectorStatistic(GetMatrixVector(cur, {{"Marginal", "Row Sum"}}), "Sum", )
                TableData[_area][_per][4][_purp] = TableData[_area][_per][5][_purp] - TableData[_area][_per][1][_purp]  //tot - intra
                
            end //_purp
        end //_area
        mat = null
    end //_per
    
    //Create daily tables
    for _area = 1 to areas.length do 
        dy_TableData = CopyArray(TableData[_area][1]) //copy first period
        for _per = 2 to periods.length do //add subsequent
            for ii = 1 to dy_TableData.length do
                for jj = 1 to dy_TableData[ii].length do
                    dy_TableData[ii][jj] = nz(dy_TableData[ii][jj]) + nz(TableData[_area][_per][ii][jj])
                end //jj
            end //ii
        end //_per
        TableData[_area] = TableData[_area] + {CopyArray(dy_TableData)}
    end //_area
    
    //Create tables for writing
    RowNames = {"Intrazonal Trips",
                "Interzonal Origins",
                "Total Origins",
                "Interzonal Destinations",
                "Total Dest."}
    ColNames = purps + {"Total"}
    
    pers_dy = periods + {"Daily"}
    Tables = null
    for _area = 1 to areas.length do 
        for _per = 1 to pers_dy.length do 
        
            //Compute marginals, but keep row totals only
            TableData[_area][_per] = Perf.Marginals(TableData[_area][_per])
            TableData[_area][_per] = Subarray(TableData[_area][_per], 1, TableData[_area][_per].length - 1)
        
            //Create table
            per = pers_dy[_per]
            TB = null
            TB.Name = per + ' Assigned Trips <span class="grey">('+areas[_area][1]+')</span>'
            TB.Section1 = areas[_area][1]
            TB.Table.TableData = TableData[_area][_per]
            TB.Table.RowNames = RowNames
            TB.Table.ColNames = ColNames
            
            Tables = Tables + {CopyArray(TB)}
        
        end //_area
    end //_per
    
    Perf.WriteTables(Tables, )
    Return(True)

EndMacro //End of Assigned Trips Summary

// ****************************************************************************************************************
// &&& Transit Assignment Summary
//
Macro "RPT Transit Assignment" (Perf)
    
    //Define intermediate files
    inrts_file = Perf.Args.Input.Routes.Value
	tdbd_file= Perf.Args.Output.INI.TrNetwork.Value
	//Define the (output copy of) route system file
	//**Special case - uses the name from the input file - needed for rts copy**
	outtmp = SplitPath(tdbd_file)
	intmp  = SplitPath(inrts_file)
	rts_file = outtmp[1]+outtmp[2]+intmp[3]+intmp[4]
	
	onoff_file = Perf.Args.Output.PST.DAYonoff.Value
				   
	pa_files = {Perf.Args.Output.MOD.TrOPW.Value, 
				Perf.Args.Output.MOD.TrPKW.Value,
				Perf.Args.Output.MOD.TrPKD.Value} 

		
NextStep = "Aggregate boarding data"
SetStatus(1, NextStep, )

    onoff_vw = OpenTable("OnOff", "FFB", {onoff_file}, )
    agg_vw = AggregateTable("OnOff_Agg", onoff_vw+"|", "MEM", null, "ROUTE", {{"On", "SUM"}}, {{"Missing as zero"}})
    CloseView(onoff_vw)
    
NextStep = "Load the Route System"
SetStatus(1, NextStep, )

	tdbd_info = GetDBInfo(tdbd_file)
    map = CreateMap("Route System", {{"Scope", tdbd_info[1]},{"Auto Project", "True"}})
    lyrs = AddRouteSystemLayer(map, "Route System", rts_file,)
    RunMacro("Set Default RS Style", lyrs, "True", "True")
    route_lyr = lyrs[1]
    stop_lyr  = lyrs[2]
	tnode_lyr = lyrs[4]
	tlink_lyr = lyrs[5]
    
NextStep = "Get On-Off Data"
SetStatus(1, NextStep, )

    //Join to routes to match ID with name
    join_vw = JoinViews("join", agg_vw+".ROUTE", route_lyr+".Route_ID", )
    
    //Load data into a table
    dim Data[2] //two columns, to be transposed before write
    Vs = GetDataVectors(join_vw+"|", {"Route_Name", "On"}, {{"Sort Order", {{"Route_Name", "Ascending"}}}})
    OnTotal = VectorStatistic(Vs[2], "Sum", )
    Data[1] = V2A(Vs[1])
    Data[2] = V2A(Vs[2])
    
    CloseView(join_vw)
    CloseView(agg_vw)
    CloseMap(map)
    
NextStep = "Get Linked Transit Trips"
SetStatus(1, NextStep, )
    
	TripTotal = 0
	for i = 1 to pa_files.length do
		mat = OpenMatrix(pa_files[i], )
		cur = CreateMatrixCurrency(mat, "Total", , , )
		marg = GetMatrixVector(cur, {{"Marginal", "Column Sum"}})
		TripTotal = TripTotal + nz(VectorStatistic(marg, "Sum", ))
		
		mat = null
		cur = null
	end
    
    RowNames = Data[1] + {"Total Boardings", "Linked Trips", "Boardings per Trip"}
    TableData = TransposeArray({Data[2] + {OnTotal, TripTotal, OnTotal/TripTotal}})
    
NextStep = "Write boarding tables"
SetStatus(1, NextStep, )

    //Separate format for last row
    dim fmts[TableData.length, TableData[1].length]
    dim CellStyles[TableData.length, TableData[1].length+1] //also include row title
    for ii = 1 to fmts.length do 
        fmts[ii][1]="*,."
    end
    fmts[fmts.length][1] = "*.00"
    for ii = 1 to CellStyles[1].length do
        CellStyles[fmts.length - 2][ii] = "font-weight: bold;" //make total 2 from bottom bold
    end
    
    TB = null
    TB.Name = "Transit Boardings Summary"
    TB.Table.TableData = TableData
    TB.Table.RowNames = RowNames
    TB.Table.ColNames = {"Boardings"}
    TB.Table.Formats = fmts
    TB.Table.Class = "dataframe no-last-col"
    TB.Table.CellStyles = CellStyles
    
    Tables = Tables + {TB}
    
    //Add a boardings chart
    CH = null
    CH.Name = "Transit Boardings by Route"
    CH.Chart.CanvasID = "transit_boardings_chart"
    CH.Chart.Type = 'bar'
    CH.Chart.Labels = Data[1]
    CH.Chart.Data = {Round(A2V(Data[2]), 0)}
    CH.Chart.Names = {"Boardings"}
    CH.Chart.XAxis = "Route"
    CH.Chart.YAxis = "Daily Boardings"
    
    
    Tables = Tables + {CH}
	
    Perf.WriteTables(Tables)
	
	SetStatus(1, "@System0", )
    Return(True)
EndMacro //End of Transit Assignment Summary

// ****************************************************************************************************************
// &&& Daily Assignment Summary
//

Macro "RPT Vehicle Assignment" (Perf) 

    shared UT

    //FT and AT information
    ft_no = UT.Values(Perf.Info.FT)  //Returns a list of numbers only
    at_no = UT.Values(Perf.Info.AT)

    //Summary area - areas = {name, query}
    areas = Perf.ActiveAreas("Network")  //Can be Network or Zones
    
    //Define files
	dbd_file = Perf.Args.Output.INI.RdNetwork.Value
	flow_file = Perf.Args.Output.PST.DailyFlow.Value //daily flows only

    
    //Define tables to hold data
    dim tVMT[areas.length]
    dim tVHT[areas.length]
    dim tDelay[areas.length]
    dim tSpeed[areas.length]
    
    //Collect tables into an array - only include tables that are to be written
    MainTables = 5 //Number of tables to show in main group (not collapsed)
    VGTable = 5 //Table by volume group - differnt headers/footers
    TableGroup = {tVMT, tVHT, tDelay, tSpeed}
    NameGroup = {"VMT - Vehicle Miles Traveled", 
                 "VHT - Vehicle Hours Traveled",
                 "Hours of Congestion Delay", 
                 "Average Congested Speed"}
                
    FormatGroup = {"*0,.", "*0,.", "*0,.", "*0,.0"}
    
    //Open dbd network
    RunMacro("TCB Add DB Layers", dbd_file,,)
    layers = RunMacro("TCB get DB line and node layers", dbd_file)
    node_lyr = layers[1]
    link_lyr = layers[2]
    
    //Open and join flows
    flow_vw = OpenTable("Flow", "FFB", {flow_file,})
    join_vw = JoinViews("Network+Flow", link_lyr+".ID", flow_vw+".ID1", )
    
    //Run cross-classification for each area and for each variable
    for _area = 1 to areas.length do
        area_name = areas[_area][1]
        area_qry = areas[_area][2]
        SetView(join_vw)
        setcount = SelectByQuery("Sel", "Several", "Select * where FT > 0", )
        setcount = SelectByQuery("Sel", "Subset", area_qry, )
        
        //Only summarize if links are selected
        if setcount > 0 then do
            //Load data from view
            {FTv, ATv, FLOWv, AB_FLOWv, BA_FLOWv, AB_TIMEv, BA_TIMEv, LENGTHv, FFv} = GetDataVectors(join_vw+"|Sel", 
            {"FT", "AT", "TOT_Flow", "AB_Flow", "BA_Flow", "AB_TIME", "BA_TIME", "Length", "FF_TIME"}, )
                               
            //Math
            VMTv = FLOWv * LENGTHv
            FFVHTv = (FLOWv * FFv) / 60
            VHTv = (nz(AB_FLOWv * AB_TIMEv) + nz(BA_FLOWv * BA_TIMEv)) / 60
            DELAYv = Max(VHTv - FFVHTv, 0) //prevent "negative zero"
           
            
            //Compute cross-class with marginals
            tVMT[_area] = Perf.CrossTab(FTv, ATv, VMTv, True)
            tVHT[_area] = Perf.CrossTab(FTv, ATv, VHTv, True)
            tDelay[_area] = Perf.CrossTab(FTv, ATv, DELAYv, True)
            
            //Compute speeds from aggregate
            dim t[tVMT[_area].length, tVMT[_area][1].length]
            tSpeed[_area] = CopyArray(t)
            for ii = 1 to tVMT[_area].length do
                for jj = 1 to tVMT[_area][ii].length do
                    if tVHT[_area][ii][jj] = 0 then tSpeed[_area][ii][jj] = 0
                    else tSpeed[_area][ii][jj] = tVMT[_area][ii][jj] / tVHT[_area][ii][jj]
                end
            end
            
        end //if records were selected
    end //end loop over summary areas
    
    //Re-organize table arrays into 
    //Tables[Area][Statistic], sort names and sub-headers too
    Tables = null
    TableNames = null
    NumberFormats = null
    SubHeaders = null
    SubHeaders2 = null
    for ii = 1 to areas.length do
        for jj = 1 to TableGroup.length do
            TB = null
            TB.Section1 = areas[ii][1]
            TB.Name = NameGroup[jj] + ' <span class="grey">('+areas[ii][1]+')</span>' 
            
            TB.Table.TableData = TableGroup[jj][ii]
            TB.Table.Formats = FormatGroup[jj]
            

            Tables = Tables + {CopyArray(TB)}
        end
        
    end
    
    Perf.WriteTables(Tables, )
    
    Return(True)
    
EndMacro

// ****************************************************************************************************************
// &&& Loaded Speed Summary

Macro "SLO Perf Speed Summary" (Args, report_file, table_file)
    //Update progres bar
    shared canned, progtot, prognum, do_sum
    prog = r2i((prognum / progtot) * 100)
    canned = UpdateProgressBar("Summarizing Loaded Speeds", prog)
    if canned = "True" then goto quit
    
    //Define page header
    section = 10
    page = 1
    title = "Loaded Speed Summary"

    //Load information
    shared ft_info, at_info, per_info
    shared area_info, area_flag
    //area_info contains queries to summarize individually
    //area_flag contains flags that define inclusion of areas
    
    //Define files
    dbd_file   =  Args.Output.INI.RdNetwork.Value
	flow_files = {Args.Output.PST.AMflow.Value, //AM    - Must match per_info!!
				  Args.Output.PST.PMflow.Value, //PM
				  Args.Output.PST.OPflow.Value, //OP
				  Args.Output.PST.DailyFlow.Value} //daily

    //create number-only ft, at index
    dim ft_no[ft_info.length]
    for i = 1 to ft_info.length do
        ft_no[i] = ft_info[i][2]
    end
    dim at_no[at_info.length]
    for i = 1 to at_info.length do
        at_no[i] = at_info[i][2]
    end
    
    //Dimension arrays to hold data: TableXXX[area][ft(row)][at(col)]
    dim SPDVMTt[area_info.length, per_info.length+1, ft_info.length, at_info.length]    
    dim VMTt   [area_info.length, per_info.length+1, ft_info.length, at_info.length]    
    dim Tables [area_info.length, per_info.length+1, ft_info.length+1, at_info.length+1]  //per_info.length+1 for FF speeds
                                                                                                        //at and ft +1 for totals
    
    //Open dbd network
    RunMacro("TCB Add DB Layers", dbd_file,,)
    layers = RunMacro("TCB get DB line and node layers", dbd_file)
    node_lyr = layers[1]
    link_lyr = layers[2]
    
    for per = 1 to per_info.length+1 do  //+1 for free-flow speeds
        //use daily if doing free-flow speeds
        if per > per_info.length then do
            flow_vw = OpenTable("Flows", "FFB", {flow_files[per-1],})
            join_vw = JoinViews("Join", flow_vw+".ID1", link_lyr+".ID", )
        end    
        //Otherwise, load directly from chosen period
        else do
            flow_vw = OpenTable("Flows", "FFB", {flow_files[per],})
            join_vw = JoinViews("Join", flow_vw+".ID1", link_lyr+".ID", )
        end

        //Loop over summary areas
        for area = 1 to area_info.length do
            SetView(join_vw)
            if area_flag[area] = 1 and area_info[area][3] = 1 then setcount = SelectByQuery("SummaryArea", "Several", area_info[area][2], )
            else setcount = 0
            if setcount > 0 then do
            
                FTv = GetDataVector(join_vw+"|SummaryArea", "FT", )
                ATv = GetDataVector(join_vw+"|SummaryArea", "AT", )
                VMTv = GetDataVector(join_vw+"|SummaryArea", "TOT_V_Dist_T", )
        
                if per > per_info.length then do //Get free-flow SpeedVMT
                    FFSPDv = GetDataVector(join_vw+"|SummaryArea", "FF_SPD", )
                    FFVMTv = GetDataVector(join_vw+"|SummaryArea", "TOT_V_Dist_T", )
                    SPDVMTv= FFSPDv * FFVMTv
                    FFSPDv = null
                    FFVMTv = null
                end
                else do  //or congested SpeedVMT
					expr = CreateExpression(join_vw, "CSPEEDVMT", "(nz(AB_V_Dist_T)*nz(AB_Speed))+ (nz(BA_V_Dist_T)*nz(BA_Speed))", )
	                SPDVMTv = GetDataVector(join_vw+"|SummaryArea", "CSPEEDVMT", )
                end

                CreateProgressBar("Aggregating Data", "False")
                for I = 1 to FTv.length do
                    prog = r2i((I / FTv.length)*100)
                    UpdateProgressBar("Aggregating Data", prog)

                    ft = ArrayPosition(ft_no, {FTv[I]}, )
                    at = ArrayPosition(at_no, {ATv[I]}, )
                    
                    //Put values in the table
                    SPDVMTt[area][per][ft][at] = nz(SPDVMTt[area][per][ft][at]) + nz(SPDVMTv[I])
                    VMTt[area][per][ft][at] = nz(VMTt[area][per][ft][at]) + nz(VMTv[I])
                end //end loop over selected records
                DestroyProgressBar()
            end //end if summary flag = 1 and setcount > 0
            SPDVMTt[area][per] = RunMacro("Array Marginals", SPDVMTt[area][per])
            VMTt[area][per] = RunMacro("Array Marginals", VMTt[area][per])

        end //end loop over summary areas
        CloseView(join_vw)
        CloseView(flow_vw)
        
        //Divide sum(SPEEDVMT) / sum(VMT)
        for area = 1 to Tables.length do
            for ft = 1 to Tables[area][per].length do
                for at = 1 to Tables[area][per][ft].length do
                    if VMTt[area][per][ft][at] > 0 then do
                        Tables[area][per][ft][at] = SPDVMTt[area][per][ft][at] / VMTt[area][per][ft][at]
                    end
                end
            end
        end //end loop over area, ft, at
        
    end //end loop over periods


	//Write tables
    fp = OpenFile(report_file, "a")
	HeaderInfo = {title, section, page}
	TableInfo = null
	wTables = null   //Use wtables for the WriteTables function - Tables already used.
	fmt = null
	for area = 1 to area_info.length do
        if area_flag[area] = 1 and area_info[area][3] = 1 then do
			for per = 1 to per_info.length+1 do
				if per > per_info.length then do
					TableInfo = TableInfo + {{"Assignment Freeflow Speed Summary", area_info[area][1]}}
					wTables = wTables + {Tables[area][per]}
					fmt = fmt + {"*0,.0"}
				end
				else do
					TableInfo = TableInfo + {{per_info[per] + " Loaded Speed Summary", area_info[area][1]}}
					wTables = wTables + {Tables[area][per]}
					fmt = fmt + {"*0,.0"}
				end
			end
		end
	end

	{ColInfo, RowInfo} = RunMacro("Create Headings", at_info, ft_info)
	
	RunMacro("WriteTables", fp, HeaderInfo, TableInfo, wTables, RowInfo, ColInfo, fmt)
    CloseFile(fp)  

    quit:
EndMacro

// ****************************************************************************************************************
// &&& Validation Summary

// *************************************************************************************************
// Assignment Validation

Macro "RPT Vehicle Validation DY" (Perf)

    RunMacro("Vehicle Validation", Perf, "DY")

EndMacro

Macro "RPT Vehicle Validation AM" (Perf)

    RunMacro("Vehicle Validation", Perf, "AM")

EndMacro

Macro "RPT Vehicle Validation PM" (Perf)

    RunMacro("Vehicle Validation", Perf, "PM")

EndMacro

Macro "Vehicle Validation" (Perf, per)

    shared UT

    //FT and AT information
    ft_no = UT.Values(Perf.Info.FT)  //Returns a list of numbers only
    at_no = UT.Values(Perf.Info.AT)

    //Summary area - areas = {name, query}
    areas = Perf.ActiveAreas("Network")  //Can be Network or Zones
    
    //Define files
    dbd_file = Perf.Args.[Highway DB]
    if per = "DY" then do
        flow_file = Perf.Args.[Hwy Day Final Flow Table] //daily flows only
        c_fld = "DAY_CNT"       //count field
    end else if per = "AM" then do
        flow_file = Perf.Args.[Hwy AM Final Flow Table] //daily flows only
        c_fld = "AM_CNT"       //count field
    end else if per = "PM" then do
        flow_file = Perf.Args.[Hwy PM Final Flow Table] //daily flows only
        c_fld = "PM_CNT"       //count field
    end
    
    //Define tables to hold data
    dim tCount[areas.length]
    dim tVolsOnC[areas.length]
    dim tCountVMT[areas.length]
    dim tModVMT[areas.length]
    dim tCountOne[areas.length]
    dim tOne[areas.length]
    dim tSquareError[areas.length]
    dim tPctSqrError[areas.length]
    
    //These are be computed based on above results
    dim tVolRatio[areas.length]
    dim tVMTRatio[areas.length]
    dim tRMSE[areas.length]
    dim tPRMSE[areas.length]
    dim tRMSEgrp[areas.length]
    dim Charts[areas.length]
    
    //Collect tables into an array - only include tables that are to be written
    MainTables = 5 //Number of tables to show in main group (not collapsed)
    VGTable = 5 //Table by volume group - differnt headers/footers
    TableGroup = {tVMTRatio, tVolRatio, tPRMSE, tRMSE, tRMSEgrp, tCountVMT, tModVMT, tCountOne, tOne}
    NameGroup = {"Modeled VMT / Count VMT", 
                "Modeled Volume / Count Volume",
                "Percent Root Mean Square Error", 
                "Root Mean Square Error",
                "RMSE by volume group",
                
                "Count VMT",
                "Modeled VMT on Links With Counts",
                "Number of Links with Counts",
                "Number of Links"}
                
    FormatGroup = {"*0.0%", "*0.0%", "*0.0%", "*0,.", "*0,.000", "*0,.", "*0,.", "*0,.", "*0,."}
    
    //RMSE Volume Group cutpoints
    VolGrps = {1000, 5000, 10000, 20000, 30000, 50000, 100000}
               
   
    //Open dbd network
    RunMacro("TCB Add DB Layers", dbd_file,,)
    layers = RunMacro("TCB get DB line and node layers", dbd_file)
    node_lyr = layers[1]
    link_lyr = layers[2]
	
    SetView(link_lyr)
	// Create formula field FT based on AB_Facility_Type
	CreateExpression(link_lyr, "FT", "if CC = 1 then 10 else (if AB_Facility_Type<100 then  floor(AB_Facility_Type/10) else 99)", ) 	// SCAG: Need to add the field "CC" in the highway network
    CreateExpression(link_lyr, "AT", "AB_AreaType", ) 	// SCAG: use AB_AreaType as AT
	
	
    //Open and join flows
    flow_vw = OpenTable("Flow", "FFB", {flow_file,})
    join_vw = JoinViews("Network+Flow", link_lyr+".ID", flow_vw+".ID1", )
    
    //Run cross-classification for each area and for each variable
    for _area = 1 to areas.length do
        area_name = areas[_area][1]
        area_qry = areas[_area][2]
        SetView(join_vw)
        setcount = SelectByQuery("Sel", "Several", "Select * where FT > 0", )
        setcount = SelectByQuery("Sel", "Subset", area_qry, )
        
        //Only summarize if links are selected
        if setcount > 0 then do
            //Load data from view
            {FTv, ATv, COUNTv, FLOWv, LENGTHv} = GetDataVectors(join_vw+"|Sel", 
            {"FT", "AT", c_fld, "TOT_Flow", "Length"}, )
                               
            //Math
            VolsOnCv = if COUNTv > 0 then FLOWv else 0
            CountVMTv = if COUNTv > 0 then (COUNTv * LENGTHv) else 0
            ModVMTv = if COUNTv > 0 then (FLOWv * LENGTHv) else 0
            CountONEv = if COUNTv > 0 then 1 else 0
            ONEv = Vector(COUNTv.Length, "Long", {{"Constant", 1}})
            SqERRv = if COUNTv > 0 then Pow(FLOWv - COUNTv, 2) else 0
            PctSqERRv = if COUNTv > 0 then SqERRv / Pow(FLOWv, 2) else 0
            
            //RMSE Volume groups
            prev_grp = 0
            VolGRPv = Vector(ONEv.length, "Long", {{"Constant", 0}})
            VolGrpLabels = null
            VolGrpFmts = null
            for _vg = 1 to VolGrps.length do
                VolGRPv = if (VolsOnCv >= prev_grp and VolsOnCv < VolGrps[_vg]) then _vg else VolGRPv
                VolGrpLabels = VolGrpLabels + {Format(prev_grp, "*0,.") + " - " + Format(VolGrps[_vg], "*0,.")}
                prev_grp = VolGrps[_vg]
                VolGrpFmts = VolGrpFmts + {{"*,.", "*,.", "*.0%"}}
            end
            //Top end group
            _vg = VolGrps.length
            VolGRPv = if (VolsOnCv >= VolGrps[_vg]) then _vg+1 else VolGRPv
            VolGrpLabels = VolGrpLabels + {Format(VolGrps[_vg], "*0,.") + " and up", "All Links"}
            
            //Formats for &up and total rows
            VolGrpFmts = VolGrpFmts + {{"*,.", "*,.", "*.0%"}}
            VolGrpFmts = VolGrpFmts + {{"*,.", "*,.", "*.0%"}}
            
            //Compute cross-class with marginals
            tCount[_area] = Perf.CrossTab(FTv, ATv, COUNTv, True)
            tVolsOnC[_area] = Perf.CrossTab(FTv, ATv, VolsOnCv, True)
            tCountVMT[_area] = Perf.CrossTab(FTv, ATv, CountVMTv, True)
            tModVMT[_area] = Perf.CrossTab(FTv, ATv, ModVMTv, True)
            tCountOne[_area] = Perf.CrossTab(FTv, ATv, CountONEv, True)
            tOne[_area] = Perf.CrossTab(FTv, ATv, ONEv, True)
            tSquareError[_area] = Perf.CrossTab(FTv, ATv, SqERRv, True)
            tPctSqrError[_area] = Perf.CrossTab(FTv, ATv, PctSqERRv, True)
            
            vgOpts = null
            vgOpts.RowList = V2A(Vector(VolGrps.length+1, "Long", {{"Sequence", 1, 1}}))
            vgOpts.ColList = {1}
            
            vgCount = Perf.CrossTab(VolGRPv, ONEv, COUNTv, True, vgOpts)
            vgCountOne = Perf.CrossTab(VolGRPv, ONEv, CountONEv, True, vgOpts)
            vgSquareError = Perf.CrossTab(VolGRPv, ONEv, SqERRv, True, vgOpts)
            vgPctSqrError = Perf.CrossTab(VolGRPv, ONEv, PctSqERRv, True, vgOpts)
            
            //Compute additional tables based on totals by FT/AT
            dim tmpVolRatio[tCount[_area].Length, tCount[_area][1].Length]
            dim tmpVMTRatio[tCount[_area].Length, tCount[_area][1].Length]
            dim tmpRMSE[tCount[_area].Length, tCount[_area][1].Length]
            dim tmpPRMSE[tCount[_area].Length, tCount[_area][1].Length]
            for row = 1 to tCount[_area].length do
                for col = 1 to tCount[_area][1].length do
                    tmpVolRatio[row][col] = tVolsOnC[_area][row][col] / zn(tCount[_area][row][col], 0.0001)
                    tmpVMTRatio[row][col] = tModVMT[_area][row][col] / zn(tCountVMT[_area][row][col], 0.0001)
                                                  
                    tmpRMSE[row][col] = sqrt(tSquareError[_area][row][col] / zn(tCountOne[_area][row][col] - 1, 0.0001))
                    tmpPRMSE[row][col] = tmpRMSE[row][col] / zn(tCount[_area][row][col] / zn(tCountOne[_area][row][col], 0.0001), 0.0001)
                    
                end //col
            end //row
            
            //Percent RMSE by volume group
            dim tmpGrpRMSE[vgCount.length, 3]
            for row = 1 to vgCount.length do
                tmpGrpRMSE[row][1] = vgCountOne[row][1]
                tmpGrpRMSE[row][2] = sqrt(vgSquareError[row][1] / zn(vgCountOne[row][1] - 1, 0.0001))
                tmpGrpRMSE[row][3] = tmpGrpRMSE[row][2] / zn(vgCount[row][1] / zn(vgCountOne[row][1], 0.0001), 0.0001)
            end
            
            //Place the results in the overall tables by area
            tVolRatio[_area] = CopyArray(tmpVolRatio)
            tVMTRatio[_area] = CopyArray(tmpVMTRatio)
            tRMSE[_area] = CopyArray(tmpRMSE)
            tPRMSE[_area] = CopyArray(tmpPRMSE)
            
            tRMSEgrp[_area] = CopyArray(tmpGrpRMSE)
            
            //Create a scatter chart
            dim data_fwy[2]
            dim data_art[2]
            dim data_oth[2]
            
            for ii = 1 to COUNTv.length do
                if COUNTv[ii] > 0 then do
                    if FTv[ii] = 1 then do
                        data_fwy[1] = data_fwy[1] + {COUNTv[ii]}
                        data_fwy[2] = data_fwy[2] + {FLOWv[ii]}
                    end else if (FTv[ii] <= 3 or (FTv[ii] BETWEEN 11 and 13)) then do
                        data_art[1] = data_art[1] + {COUNTv[ii]}
                        data_art[2] = data_art[2] + {FLOWv[ii]}
                    end else do
                        data_oth[1] = data_oth[1] + {COUNTv[ii]}
                        data_oth[2] = data_oth[2] + {FLOWv[ii]}
                    end
                end //have count
            end
            
            Chart = null
            Chart.CanvasID = "scatterchart_" + area_name + per
            Chart.Type = 'scatter'
            Chart.Names = {"Freeway", "Arterial", "Other"}
            Chart.Data = {data_fwy, data_art, data_oth}
            Chart.XAxis = "Count"
            Chart.YAxis = "Model Volume"
            Chart.Width = 500
            Chart.Height = 500
            
            Charts[_area] = CopyArray(Chart)
            
        end //if records were selected
    end //end loop over summary areas
    
    //Re-organize table arrays into 
    //Tables[Area][Statistic], sort names and sub-headers too
    Tables = null
    TableNames = null
    NumberFormats = null
    SubHeaders = null
    SubHeaders2 = null
    for ii = 1 to areas.length do
        for jj = 1 to TableGroup.length do
            TB = null
            TB.Section1 = areas[ii][1]
            if jj > MainTables then TB.Section2 = "Additional Details"
            TB.Name = NameGroup[jj]
            
            TB.Table.TableData = TableGroup[jj][ii]
            TB.Table.Formats = FormatGroup[jj]
            
            if jj = VGTable then do
                TB.Table.RowNames = VolGrpLabels
                TB.Table.ColNames = {"Volune Group", "Links", "RMSE", "% RMSE"}
                TB.Table.Formats = VolGrpFmts //override default with cell-specific array
            end

            Tables = Tables + {CopyArray(TB)}
        end
        
        //Scatter chart
        CH = null
        CH.Section1 = areas[ii][1]
        CH.Name = 'Scatter Plot' + ' <span class="grey">('+areas[ii][1]+')</span>' 
        CH.Chart = CopyArray(Charts[ii])
        Tables = Tables + {CopyArray(CH)}
    end
    
    Perf.WriteTables(Tables, )
    
    Return(True)
    
EndMacro


// >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
// Other usefull macros
// >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


// ****************************************************************************************************************
// Write page header
// ****************************************************************************************************************

Macro "Page Header" (fp, title, section, page)
	shared scen_name
	
	shared contents_info
	{ContentsName, step} = contents_info
	
	pad = 115 - len(title)
	
	WriteLine(fp, "<hr class=\"first\">")
	WriteLine(fp, "LSA Associates, Inc. - San Luis Obispo Citywide Travel Model - " + scen_name + "<br>")
	WriteLine(fp,  title + lpad("Page "+i2s(section)+"."+i2s(page),pad))
	WriteLine(fp, "<hr>")
	//Place anchor on the first page of each section
	if page = 1 then do
		WriteLine(fp,"<a name = \"" + ContentsName[section+1][2] + "\"></a>")
	end
EndMacro

// >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
// Macro WriteTables
// Write tables passed to the macro to the specified file.
//	fp: file pointer for spedcified open text file
//	HeaderInfo: {title, section, page} info for the page header
//	TableInfo: {"Table Name", "Table Name", ...} Table names
//	Tables: Array filled with 2D table arrays.
//	RowInfo: {RowText, RowTextWidth} Specifies text to start each row, width of row header column.
//			 Note: the length of RowText must match the number of rows in all tables.
//	ColInfo: {ColText, ColTextWidth} Specifies text written as column header, width of each column.
//	fmt: Array of number format strings for use with each table
//     -----> If the format string for a table is "x" (case sensitive) then the table will not be written!
//     -----> fmt can be a single string, allowing a standard format for all tables
//  TablesPerPage (Optional): Number of tables to be included on each page (default = 2)
// ****************************************************************************************************************
// Get row and column sums of an n x m array and add to the array.
// returns an (n+1) x (m+1) array.
// ****************************************************************************************************************

Macro "Array Marginals" (array)
    n = array.length
    m = array[1].length

    //Get row totals
    for i = 1 to n do
        if array[i].length <> m then Return({0}) //verify consistency of m dimension
        s = 0
        for j = 1 to m do
            s = s + nz(array[i][j])
        end
        array[i] = array[i] + {s}
    end
    
    //Get column totals
    dim s[m+1]
    for i = 1 to n do
        for j = 1 to m+1 do
            s[j] = nz(s[j]) + nz(array[i][j])
        end
    end
    array = array + {s}
    
    //Return updated array
    return(array)
EndMacro

// >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
// LSA Crosstab Macro
//
// This macro returns a 2 dimensional array containing cross tabluated sum values.
// INPUTS:
//	view: The view name and selection set from which to gather data
//	Fields: An array of 2 strings, {Row Field, Column Field"
//	Headers: An array defining values to include ({Value, Value, ...}, {Value, Value, ...}}
//			 Data type *must* match the type in the dataview.
//	Val: The field name to be summed.  Missing vlues are treated as zeros.

Macro "LSA Crosstab" (viewset, Fields, Headers, Val)

	dim retarr[Headers[1].Length, Headers[2].Length]
	//Step 1: Get data vectors from the viewset
	RowField = GetDataVector(viewset, Fields[1], )
	ColField = GetDataVector(viewset, Fields[2], )
	ValField = GetDataVector(viewset, Val, )
	
	//Step 2: Loop over value to place data in the correct slots
	for I = 1 to RowField.Length do
		//Get number-only indices
		row = ArrayPosition(Headers[1], {RowField[I]}, )
		col = ArrayPosition(Headers[2], {ColField[I]}, )
		
		//Add value to correct location in 2D array
		if row <> 0 and col <> 0 then 
			retarr[row][col] = nz(retarr[row][col]) + nz(ValField[I])

	end //end loop over all records
	
Return(retarr)
EndMacro

// ****************************************************************************************************************
// Creat Row and Column info arrays based on at_info and ft_info
// ****************************************************************************************************************

Macro "Create Headings" (at_info, ft_info)

	ColText = null
	for i = 1 to at_info.length do
		ColText = ColText + {at_info[i][1]}
	end
	ColText = ColText + {"Total"}
    RowText = null
	for i = 1 to ft_info.length do
		RowText = RowText + {ft_info[i][1]}
	end
	RowText = RowText + {"Total"}
	ColInfo = {ColText, 100}
	RowInfo = {RowText, 150}
    
    return({ColInfo, RowInfo})
EndMacro

Macro "CreateDSN" (Filename)
// This macro creates a DSN file that can be used to read or write data from
// the MS Access database that is passed in the variable Filename.  The returned
// value is the name of a the created DSN file, which is created in the TransCAD
// temporary directory.

    //Re-using DSN files results in incorrect management of database access.
	//Instead, a new temp DSN file is created each time this macro is run.
	//file_dir = SplitPath(Filename)
	//file_dir = file_dir[1]+file_dir[2]
	//dsn_file = file_dir + "Access.dsn"
	
	dsn_file = GetTempFilename(".dsn")


	fp = OpenFile(dsn_file, "w")
	WriteLine(fp, "[ODBC]")
	//WriteLine(fp, "DRIVER=Microsoft Access Driver (*.mdb)")
    WriteLine(fp, "DRIVER=Microsoft Access Driver (*.mdb, *.accdb)")
	WriteLine(fp, "UID=admin")
	WriteLine(fp, "UserCommitSync=Yes")
	WriteLine(fp, "Threads=3")
	WriteLine(fp, "SafeTransactions=0")
	WriteLine(fp, "PageTimeout=5")
	WriteLine(fp, "MaxScanRows=8")
	WriteLine(fp, "MaxBufferSize=2048")
//	WriteLine(fp, "FIL=MS Access")
//	WriteLine(fp, "DriverId=25")
	WriteLine(fp, "DefaultDir=" + file_dir)
	WriteLine(fp, "DBQ=" + Filename)
	CloseFile(fp)

	return(dsn_file)
EndMacro

//Macro to open an Access table using a DSN file, the table/query's name, and an
//optional unique field (e.g., ID). This macro opens the table and then saves it 
//as a fixed format binary file.
// - If target_file is specified, this macro saves the table there, overwriting 
//   any existing file. The target filename must be a ".bin" filename.
// - If target_file is not specified, this macro saves the table in a temp file
//   that will be deleted when the TransCAD program is next closed. The temp 
//   filename is not returned from this macro.  The temp filename can be
//   identified by passing a null variable as the target file with an & (e.g. 
//   &myfile)
// - This macro opens the new table and returns the view name.
//
//                                *****  ***********
Macro "OpenDSN" (dsn_name, tname, index, target_file)

	//Create a progress bar
	CreateProgressBar("Loading Data from Access - "+tname, "False")

	//Do some input data validation
	if GetFileInfo(dsn_name) = null then return()
	if target_file = null then target_file = GetTempFileName(".bin")
	tmp = SplitPath(target_file)
	if tmp[4] <> ".bin" then return()
	
	//Open the table and export to the new file
	on notfound do
		ShowMessage("Cannot open " + tname)
		on Error default
		return()
	end
	vw_mdb = OpenTable(tname+"MDB", "ODBC", {dsn_name, tname, index, })
	on notfound default
	ExportView(vw_mdb+"|", "FFB", target_file, ,)
	CloseView(vw_mdb)
	vw = OpenTable(tname, "FFB", {target_file, })
	
	//Return the new view
	DestroyProgressBar()
	Return(vw)
EndMacro

