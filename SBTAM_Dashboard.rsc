//******************************************************************************
//**                                                                          **
//**                             Dashboard                                    **
//**                                                                          **
//**          The Dashboard provides access to model output data              **
//**                                                                          **
//**               Version 2.0 - Designed for TransCAD 7                      **
//**                                                                          **
//**                                                                          **
//** ------------------------------------------------------------------------ **
//**                                                                          **
//** File Contents:                                                           **
//**  - Dashboard: Dashboard dialog box                                       **
//**  - Mapper: Flexible Map Creation                                         **
//**     Mapper dialog boxes                                                  **
//**     --> SelectCompare: Ask for a scenario to compare to                  **
//**     --> SelectMapSettings: Select link map settings                      **
//**                                                                          **
//** ------------------------------------------------------------------------ **
//**                                                                          **
//**  Many Options and defaults are currently set in the Mapper init step,    **
//**  but may be moved to an external settings file in future versions        **
//**                                                                          **
//**  Map options are currently set in the Dashboard init step                **
//**                                                                          **
//**  Performance reporting is currently managed in a separate rsc file, but  **
//**  may be moved into the dashboard rsc code at a later time                **
//**                                                                          **
//******************************************************************************

Macro "SCAG_Subregion Model Version" 
    project_version_number = 20170809
    required_tc_build   = 9090
    required_tc_version = 6.0
    return({project_vsersion_number, required_tc_build, required_tc_version}) 
EndMacro



//// *****************************************************************************
//// CreateMaps: Interactive mapping dialog box
//Dbox "Dashboard" (Args, callers, logo_bmp) Location: mapbox_x, mapbox_y
//// Array Args = Scenario (may be changed by this dialog box!)
//// Array callers = 2D array of calling dialog boxes. Each element is an array 
////                 of {ui, dbox_name} where ui is the name of a UI interface
////                 file or null for the current interface
//// String logo_bmp = Name of bitmap logo filename (232 x 71)
////                
//// *****************************************************************************
//    toolbox
	
DBox "SCAG Dashboard"
    center, center, 42, 35 toolbox NoKeyboard
    title: "SCAG Dashboard"	

    init do
    //StartMethod
        shared UT
        shared ScenArr, ScenFlag
		
		//Start of SCAG revisions
		Shared ScenArr, ScenSel, ModelInfo, StageInfo, MacroInfo, ScenName, ui_file
		Shared ScenFlag, Args, scenario_list, DashScen, logo_bmp

		ui_file = GetInterface()
		dbox_name = "SCAG Dashboard"
		model_title = "SCAG_Subregion Model"
		stages = 1
		logo_bmp = "C:\\Program Files\\TransCAD 6.0\\bmp\\SBTAM_Dashboard_logo.bmp"
		
		{ModelInfo, StageInfo, MacroInfo,} = RunMacro("TCP Load Model", model_title)
		model_table = ModelInfo[1]
		
		dim scenario_list[ScenArr.length]
		
		for i = 1 to ScenArr.length do 
			scenario_list[i] = ScenArr[i][1]
		end
		//End of SCAG revisions	
		
        
        //Set position
        static mapbox_x, mapbox_y
        if mapbox_x = null then mapbox_x = -3
        
        //Hide calling dialog boxes
        for i = 1 to callers.length do
            if callers[i][1] != null then 
                SetAlternateInterface(callers[i][1])
            //on NotFound goto next1
            HideDbox(callers[i][2])
            next1:
            on NotFound default
            if callers[i][1] != null then 
                SetAlternateInterface()            
        end

        //Default Map Scope - Use GetMapWindowScope() to get values based on the
        //                    current map window.  Y = GetMapWindowScope(ActiveNetwork)
                                                   //z = ShowArray({Y})
        def_scope = Scope(Coord(-116089503, 34634558), 233.095864, 174.282778, 0)
        //Use the following to change:
        /*
        scp = GetMapWindowScope()
        line = "Scope(Coord(" + 
            string(scp.center.lon) + ", " + string(scp.center.lat) + "), " + 
            string(scp.width) + ", " + string(scp.height) + ", 0)"
        CopyToClipboard({{"Text", line}})
        
        */
        
        //Model-specific files and parameters
        cnt_fld = "FIN_CNT"
		ft_fld = "AB_Facility_Type"
		conn_qry = "Select * Where "+ft_fld+" = 100"		// SCAG Tier 1 Centroid Connectors
        local_qry = "Select * Where "+ft_fld+" = 70"			// Minor Collector
		hide_qry = "Select * Where "+ft_fld+" = 200 or "+ft_fld+" = 400 or "+ft_fld+" = 999"	// SCAG Tier 2 Centroid Connectors or transit links 
        exp_priority = "(if nz("+cnt_fld+") > 1 then (9000 - 100*"+ft_fld+") + Length else " +
                       "(1000 - 100*"+ft_fld+") + Length)"
        exp_validation = 'if TOT_Flow > 0 then ((Format(TOT_Flow/1000, "*.") + ' + 
                         'if '+cnt_fld+' > 0 then "(" + ' + 
                         'Format('+cnt_fld+'/1000, "*.") + ")"))'
        

        //Set scenario-specific variables
        //(Including some mapper properties)
        RunMacro("SetScenario")
                     
        //List of map names (indexed by radio button type)
        //WARNING: Names must match the "TrafficCreate" button
        TrafficMap = null
        TrafficMap.MapNames = {"Validation Map",         
                               "Volume Map",             
                              // "LOS Map",                
                               "Traffic Comparison Map", 
                               "VC Map"}                 

        //List of available settings checkboxes
        //The update macro enables all of these checkboxes unless listed in
        //TrafficMap.Disable (in which case the checkbox is disabled)
        //TrafficMap.Settings = {"NCHRP", "Volumes", "Highlight", "Connectors", "Big Labels", 
        //                       "Label Connectors", "Period"}
        TrafficMap.Settings = {"Volumes", "Highlight", "Connectors", "Big Labels", 
                               "Label Connectors", "Period"}							   
        
        //List of settings to disable for each map.  If present in this array, 
        //  settings buttons will be disabled and set to the specified value.
        //  Specified value cannot be null (but can be 0)
        //TrafficMap.Disable.[Validation Map].NCHRP = 0
        TrafficMap.Disable.[Validation Map].Volumes = 1
        TrafficMap.Disable.[Validation Map].Period = 1
        
        //TrafficMap.Disable.[Volume Map].NCHRP = 0
		TrafficMap.Disable.[Volume Map].Volumes = 1
        TrafficMap.Disable.[Volume Map].Highlight = 0
        
        //TrafficMap.Disable.[LOS Map].Highlight = 0
        ////TrafficMap.Disable.[LOS Map].Period = 1
        //TrafficMap.Disable.[LOS Map].NCHRP = 0
        
        //TrafficMap.Disable.[Select Link/Node Map].NCHRP = 0
        TrafficMap.Disable.[Select Link/Node Map].Highlight = 0
        
        //TrafficMap.Disable.[Traffic Comparison Map].NCHRP = 0
        TrafficMap.Disable.[Traffic Comparison Map].Highlight = 0
        TrafficMap.Disable.[Traffic Comparison Map].Period = 1
        
        TrafficMap.Disable.[VC Map].Highlight = 0
                     
        //Default map type and options
        TrafficMap.Type = 2
       // TrafficMap.Opts.NCHRP = 1
        TrafficMap.Opts.Volumes = 0
        TrafficMap.Opts.Connectors = 1
		TrafficMap.Opts.[Big Labels] = 1
        TrafficMap.Opts.Highlight = 1
        TrafficMap.Opts.[Label Connectors] = 1
		TrafficMap.Opts.[Input Network] = 0
		TrafficMap.Opts.Locals = 1
        TrafficMap.Opts.Period = 1

        //Saved options (default to the same as original, )
        TrafficMap.Save = CopyArray(TrafficMap.Opts)
        
        //Separate SelectMap Options (just need to set defaults)
        shared SelectMap
        SelectMap = null
        SelectMap.Type = 1
        SelectMap.Opts.Period = 1
        SelectMap.Opts.Connectors = 1
        SelectMap.Opts.Labels = 1
        SelectMap.Opts.Locals = 1
        
        //Get scenario list and set variable
        //scen_list = UT.Keys(ScenArr)
        //scen_active = ScenFlag[1]
        
        
        RunMacro("Update")
		
        UT = null
		UT = CreateObject("Utilities")
        self.UT = UT		
        
    EndItem
    //EndMethod
    
    update do
        if SelectMap.MP = null then EnableItem("SelectCreate")
        else DisableItem("SelectCreate")
    enditem
    
    //Update the dialog box to enable and disable items as needed for each map
    Macro "Update" do
    
        //Enable and disable options
        MapDisable = TrafficMap.Disable.(TrafficMap.MapNames[TrafficMap.Type])
    
        for i = 1 to TrafficMap.Settings.Length do
            s = TrafficMap.Settings[i]
            
            if MapDisable.(s) != null then do 
                DisableItem(s)
                TrafficMap.Opts.(s) = MapDisable.(s)
            end
            else do
                EnableItem(s)
                TrafficMap.Opts.(s) = TrafficMap.Save.(s)
            end
        end
        
        ////Disable NCHRP if period > 1 (only allowed for daily)
        //if TrafficMap.Opts.Period > 1 then do
        //    DisableItem("NCHRP")
        //    TrafficMap.Opts.NCHRP = 0
        //end
		
    
    EndItem
    //EndMethod
    
    //Set scenario-specific variables
    Macro "SetScenario" do
    
        //Identify network and rts files
        //taz_file = Args.Input.TAZ.Value
		//indbd_file = Args.Input.Network.Value
        //dbd_file = Args.Output.INI.RdNetwork.Value
		//NetYear = Args.Param.INI.NetYear.Value
        
        taz_file = Args.TAZ_DB
		indbd_file = Args.[Highway Master DB]		//input highway network
        dbd_file = Args.[Highway DB]				//output highway network
		//NetYear = Args.Param.INI.NetYear.Value	//not available for SCAG
		
        //List of periods and corresponding flow files
		//per_list = {"Daily", "AM Period", "PM Period", "MD Period", "EVE Peak Hour", "NT Peak Hour"}
        per_list = {"Daily", "AM (6am to 9am)", "PM (3pm to 7pm)", "MD (9am to 3pm)", "EVE (7pm to 9pm)", "NT (9pm to 6am)"}
        per_abbr = {"DY", "AM", "PM", "MD", "EVE", "NT"}
        //flow_list = {Args.Output.PST.DailyFlow.Value,
        //             Args.Output.ASN.AMFlow.Value, 
        //             Args.Output.ASN.PMFlow.Value,
        //             Args.Output.ASN.OPFlow.Value}
        flow_list = {Args.[Hwy Day Final Flow Table],
                     Args.[Hwy AM Final Flow Table], 
                     Args.[Hwy PM Final Flow Table],
                     Args.[Hwy MD Final Flow Table],
					 Args.[Hwy EVE Final Flow Table],
                     Args.[Hwy NT Final Flow Table]}					 
    
        //Get select link/zone availability and query names
        //qry_file = Args.Input.SelQry.Value
		qry_file = null								//not available for SCAG
        if GetFileInfo(qry_file) != null then do
            sel_list = UT.SelectList(qry_file)
            sel_sum_file = Args.Output.PST.SelectSummary.Value	//not available for SCAG
            sel_val = 1
            qry_avail = True
        end else do
            qry_avail = False
        end
    
    enditem
    //EndMethod
    
    //StartMethod //Dialog box header (Logo, and scenario drop-down)
    //**************************************************************************
    // Model Logo
	Sample "Model Logo" 1, .5, 39, 5 transparent 
                                     contents: SamplePoint("Color Bitmap", 
                                     logo_bmp, -1, , )
    
    //popdown menu "Scenario" 1, 6, 40, 10 list: scen_list variable: scen_active do
    //
    //    //Update Args to reference the current scenario
    //    ScenFlag = {scen_active}
    //    ScenName = scen_list[scen_active]
    //    Args = ScenArr.(ScenName)
    //    
    //    //Update mapper and dialog box variables with the current scenario
    //    RunMacro("SetScenario")
    //    
    //    //Update all calling dialog boxes
    //    //(This is necessary when the senario changes so that the correct
    //    // scenario will be shown when returning to the main dbox)
    //    for i = 1 to callers.length do
    //        SetAlternateInterface(callers[i][1])
    //        UpdateDbox(callers[i][2])
    //        SetAlternateInterface()
    //    end
    //enditem

	//Start of SCAG revisions
	popdown menu "Scenario" 1, 6, 40, 10 key: alt_s list: scenario_list variable: DashScen
										help: "Select senarios"				do
        
		//Update Args to reference the current scenario
        ScenFlag = {DashScen}
        ScenName = scenario_list[DashScen]
        // Args = ScenArr.(ScenName)
		Args = RunMacro("TCP Convert to Argument Options", ScenArr[DashScen][5])	
		Args.Info.Name = ScenName
		Args.Info.ModelDir = ScenArr[DashScen][3]
		Args.Info.Description = ScenArr[DashScen][4]

		//Update mapper and dialog box variables with the current scenario
		RunMacro("SetScenario")		
		
	enditem 
	// End of SCAG revisions
	
    
    //EndMethod
    
    //StartMethod //Close buttons
    //**************************************************************************
    close do
        ExitType = "Close"
        RunMacro("Quit")
    enditem
    
    ////button "Return" 2, 28, 38, 2 prompt: "<< Return to dialog box" do
    //button "Return" 2, 29, 38, 2 prompt: "<< Return to dialog box" do
    //    ExitType = "Show"
    //    RunMacro("Quit")
    //enditem
    
    //button "Exit" same, after, 38, 2 prompt: "Exit" do
	button "Exit" 2, 30, 38, 2 prompt: "Exit" do
        ExitType = "Close"
        RunMacro("Quit")
    enditem
    
    macro "Quit" do
        for i = 1 to callers.length do
            if callers[i][1] != null then 
                SetAlternateInterface(callers[i][1])
            
            //on NotFound goto next1
            if ExitType = "Show" then ShowDbox(callers[i][2])
            else CloseDbox(callers[i][2])
            next1:
            on NotFound default
            
            if callers[i][1] != null then 
                SetAlternateInterface()            
                
        end
        Return()
    enditem
    //EndMethod
    
    
    tab list "DashTabs" 1, 8, 40, 26
    
    tab "RdMaps" prompt: "Roadway"
    
    
    //StartMethod //Traffic Map type selection
    //Traffic Maps

    radio list "Traffic Maps" 0.5, 1, 38, 15 variable: TrafficMap.Type prompt: "Traffic Maps"
    
    radio button "Validation" 2, 2.5 do RunMacro("Update") enditem         //1
    radio button "Volume" same, after do RunMacro("Update") enditem        //2
    //radio button "LOS" same, after do RunMacro("Update") enditem           //3
    radio button "Traffic Comparison" 18, 2.5 do RunMacro("Update") enditem  //4
    radio button "Volume/Capacity" same, after do RunMacro("Update") enditem //5
    
    
    //EndMethod
    //StartMethod //Traffic Map options
    
    //checkbox "NCHRP" 2, 6 
    //         variable: TrafficMap.Opts.NCHRP 
    //         help: "Create map using NCHRP-255 Adjusted daily volumes" do
    //    TrafficMap.Save.NCHRP = TrafficMap.Opts.NCHRP 
    //enditem
    //checkbox "Volumes" same, after 
	checkbox "Volumes" 2, 7
             variable: TrafficMap.Opts.Volumes 
             help: "Include volume labels on the map" do
        TrafficMap.Save.Volumes  = TrafficMap.Opts.Volumes 
    enditem
    checkbox "Connectors" same, after 
             variable: TrafficMap.Opts.Connectors 
             help: "Show centroid connectors on the map" do
        TrafficMap.Save.Connectors = TrafficMap.Opts.Connectors 
    enditem
    checkbox "Big Labels" same, after 
             variable: TrafficMap.Opts.[Big Labels]
             help: "Use larger labels for on-screen viewing" do
        TrafficMap.Save.[Big Labels] = TrafficMap.Opts.[Big Labels]
    enditem
	
    // checkbox "Highlight" 18, 6 
	checkbox "Highlight" same, after 
             variable: TrafficMap.Opts.Highlight 
             help: "Highlight high/low volumes" do
        TrafficMap.Save.Highlight = TrafficMap.Opts.Highlight 
    enditem
    //checkbox "Label Connectors" same, after 
	checkbox "Label Connectors" 18, 7
             variable: TrafficMap.Opts.[Label Connectors]
             help: "Show labels oncentroid connectors" do
        TrafficMap.Save.[Label Connectors] = TrafficMap.Opts.[Label Connectors]
    enditem
    checkbox "Input Network" same, after 
             variable: TrafficMap.Opts.[Input Network]
             help: "Create the map using the input geographic file" do
        TrafficMap.Save.[Input Network] = TrafficMap.Opts.[Input Network]
    enditem	
    checkbox "Show Locals" same, after 
             variable: TrafficMap.Opts.Locals
             help: "Show local streets on the map" do
        TrafficMap.Save.Locals = TrafficMap.Opts.Locals
    enditem	
	

    
    popdown menu "Period" 20, 13, 16, 7 List: per_list 
                 variable: TrafficMap.Opts.Period
                 help: "Traffic assignment time period" do
        TrafficMap.Save.Period = TrafficMap.Opts.Period
		RunMacro("Update")
    enditem
                 
    
    //EndMethod
    //StartMethod //Traffic Map Create button 
    button "TrafficCreate" 2, 13, 15, 1.5 prompt: "Create" do
        HideDbox()
        
        //If using NCHRP adjustment, set adj variable
        if TrafficMap.Opts.NCHRP then adj = "_NCHRP"
        else adj = ""
        
        //Initialize mapper object
        MP = CreateObject("Mapper")
        
        MapName = TrafficMap.MapNames[TrafficMap.Type]

        //If running a comparison, ask for a scenario
        if MapName = "Traffic Comparison Map" then do
            comp_flow = MP.GetScenario()
            if comp_flow = null then goto nomap
        end
        
        //Initialize map
        MP.Files.Zones = taz_file
        MP.Scope = def_scope
        MP.[Count Field] = cnt_fld
        MP.MapName = MapName
        if TrafficMap.Opts.[Input Network] then do
			MP.Create(indbd_file, True) //True= don't redraw map
			MP.Settings.Network.FTField = ft_fld+"_" + NetYear
		end
		else MP.Create(dbd_file, True) //True= don't redraw map
        
        //Select centroid connectors (optionally show them)
        MP.Connectors(conn_qry, TrafficMap.Opts.Connectors)
        //Select local streets (hide if not selected)
        if hide_qry != null then do
            MP.Select("Local Streets", local_qry, TrafficMap.Opts.Locals)
        end
        //Select links to hide (always hide them)
        MP.Select("Inactive Links", hide_qry, False, {{"LineType", {MP.Styles("Dash"), MP.Colors("Red"), 1.5}}})

        //Identify and join the flow file
        flow_file = flow_list[TrafficMap.Opts.Period]
        MP.JoinFlows(flow_file, True)
        
        //*** Replace link themes ***
        
        //Add bandwidth and LOS theme if enabled
        if MapName = "LOS Map" then do
            MP.ClearThemes(MP.Layers.Links)
            MP.Bandwidths(MP.Views.NetFlow+".TOT_Flow"+adj, 
                          {{"Data Source", "Screen"}})
            MP.LOS("LOS_MAP"+adj)
        end
        //Or, add a traffic comparison theme
        else if MapName = "Traffic Comparison Map" then do 
            MP.ClearThemes(MP.Layers.Links)
            MP.CompareFlows(comp_flow, 5) 
        end
        //V/C map
        else if MapName = "VC Map" then do
            MP.ClearThemes(MP.Layers.Links)
            MP.Bandwidths(MP.Views.NetFlow+".TOT_Flow", 
                          {{"Data Source", "All"}, 
                           {"Line Style", "Solid"}})
            
                             
            CreateExpression(MP.Views.NetFlow, "AB_MVOC", 
                             "AB_VOC", )
            CreateExpression(MP.Views.NetFlow, "BA_MVOC", 
                             "BA_VOC", )
            CreateExpression(MP.Views.NetFlow, "TOT_MVOC", 
                             "Max(nz(AB_MVOC), nz(BA_MVOC))", )
                          
            //!!! Directional MP.VOC("AB_MVOC")
            MP.VOC("TOT_MVOC")
            
            SetLineColor(MP.Layers.Links+"|CentroidConnectors", 
                         MP.Colors("LtGray"))
        end
        else do
            //Default to the FT Theme, but use the default FT field (FT)
            //rather that a year-based FT theme
            MP.FTTheme()
        end
        
        //Add link labels if active
        Opts = null
        Opts.[Priority Expression] = exp_priority
        if !TrafficMap.Opts.[Label Connectors] then Opts.CC.Expression = 'null'
		if TrafficMap.Opts.[Big Labels] then do
			Opts.Font = "Arial|12"
			Opts.CC.Font = "Arial|10"
		end
        if MapName = "Validation Map" then do
            Opts.Highlight = TrafficMap.Opts.Highlight
            Opts.Suffix = "validation"+per_abbr[TrafficMap.Opts.Period]
            exp = exp_validation
            MP.Label(exp, Opts)
        end
        else if MapName = "Traffic Comparison Map" then do
            Opts.ExpressionView = MP.Views.NetCompareFlow
            Opts.Suffix = "diff"+per_abbr[TrafficMap.Opts.Period]
            MP.Label('if DIFFCOLOR = 0 then null ' +
                     'else Format(ABSDIFF/1000, "*.0")', Opts)
        end
        else if TrafficMap.Opts.Volumes then do //Non-validation
			if TrafficMap.Opts.Period = 1 then 
				exp = 'if TOT_Flow > 0 then Format(TOT_Flow'+adj+'/1000, "*.")'
			else 
				exp = 'if TOT_Flow > 0 then Format(TOT_Flow'+adj+'/1000, "*.0")'
            Opts.Suffix = "volume"+per_abbr[TrafficMap.Opts.Period]
            MP.Label(exp, Opts)
        end
			
		//Make read only if not input network
		CurrentViews = GetViews()
		CurrentViews = CurrentViews[1]
		if !TrafficMap.Opts.[Input Network] then do
			for _lyr = 1 to MP.Layers.length do
				lyr = MP.Layers[_lyr][2]
				if ArrayPosition(CurrentViews, {lyr}, ) > 0 then
					SetViewReadOnly(lyr, "True")
			end
			for _vw = 1 to MP.Views.length do
				vw = MP.Views[_vw][2]
				if ArrayPosition(CurrentViews, {vw}, ) > 0 then
					SetViewReadOnly(vw, "True")
			end
		
		end

        MP.Redraw()
        SetMapRedraw(MP.Map, "True")
        //If cancelled, we will skip to here
        nomap:
        MP = null //release mapper object 
        
        ShowDbox()
    enditem
    //EndMethod
    
    //StartMethod //Select Link/Zone Map Selection
    tab "SelectMaps" prompt: "Select Link/Zone"
    
    popdown menu "SelQry" 13, 2 list: sel_list variable: sel_val prompt: "Select Query:" do RunMacro("UpdateSelect") enditem
    popdown menu "SelPer" same, after list: per_list variable: SelectMap.Opts.Period prompt: "Period:" do RunMacro("UpdateSelect") enditem
        
    radio list "SelectType" 0.5, 5, 25, 6 variable: SelectMap.Type prompt: "Display Type"
        
    radio button "Link" 2, 6.5 do RunMacro("UpdateSelect") enditem                //1
    radio button "Origins" same, after  do RunMacro("UpdateSelect") enditem        //2
    radio button "Destinations" same, after  do RunMacro("UpdateSelect") enditem   //3
    //radio button "O and D" same, after  do RunMacro("UpdateSelect") enditem        //4 !!! TODO
    
    checkbox "Connectors" 2, 11.5 variable: SelectMap.Opts.Connectors do RunMacro("UpdateSelect") enditem
    checkbox "Labels" same, after variable: SelectMap.Opts.Labels do RunMacro("UpdateSelect") enditem
    checkbox "Show Locals" 18, 11.5 variable: SelectMap.Opts.Locals do RunMacro("UpdateSelect") enditem
    //EndMethod - Select link/zonem map selection
    
    //StartMethod //Select Link/Zone Map Create button 
    button "SelectCreate" 2, 15, 15, 1.5 prompt: "Create" do
        HideDbox()
        //
        // ** POTENTIAL ADDITIONS **
        //   * Add color pickers for links and zones
        //   * Add ability to combine Link+Zone?
        //   * Add ability to map Os and Ds together with dot density?
        //   * Add TAZ labels?
        
        //Initialize mapper object
        MP = CreateObject("Mapper")
        SelectMap.MP = MP
        
        MapName = "Select Link/Zone Map"
        sel_qry = sel_list[sel_val]
        
        //Initialize map
        MP.Files.Zones = taz_file
        MP.Scope = def_scope
        MP.[Count Field] = cnt_fld
        MP.MapName = MapName
        Opts.[Close macro] = "CloseSelect"
        if TrafficMap.Opts.[Input Network] then do
			MP.Create(indbd_file, True) //True= don't redraw map
			MP.Settings.Network.FTField = ft_fld+"_" + NetYear
		end
		else MP.Create(dbd_file, True, Opts) //True= don't redraw map
        
        //Select centroid connectors (optionally show them)
        MP.Connectors(conn_qry, SelectMap.Opts.Connectors)
        if hide_qry != null then do
            MP.Select("Local Streets", local_qry, SelectMap.Opts.Locals)
        end
        //Select links to hide (always hide them)
        MP.Select("Inactive Links", hide_qry, False, {{"LineType", {MP.Styles("Dash"), MP.Colors("Red"), 1.5}}})

        //Join select volumes for all periods
        MP.JoinSelect(flow_list, per_abbr, sel_list)
        
        //Join the OD summary
        MP.JoinData(sel_sum_file, "TAZ")
        
        //*** Create/Refresh the FT Theme ***
        ft_theme = MP.FTTheme({{"CreateOnly", True}})
        
        
        //*** Create Bandwidth themes for diffferent options ***
        dim link_themes[per_abbr.length, sel_list.length]
        for _per = 1 to per_abbr.length do
            per = per_abbr[_per]
            for _sel = 1 to sel_list.length do
                Opts = null
                Opts.[Data Source] = "Screen"
                Opts.ThemeName = per + " " + sel_list[_sel] + " Flow"
                Opts.CreateOnly = True
                link_themes[_per][_sel] = MP.Bandwidths(MP.Views.SelectFlows+'.AB_Flow_'+sel_list[_sel]+"_"+per, Opts)
            end
        end
        
        //Set up link labels
        SelLabelOpts = null
        SelLabelOpts.[Priority Expression] = exp_priority
        SelLabelOpts.Font = "Arial|12"
        SelLabelOpts.CC.Font = "Arial|10"
        SelLabelOpts.ExpressionView = MP.Views.SelectFlows
        SelLabelOpts.ClearOld = True
        SelLabelOpts.Suffix = "select"
        dim sel_exp[per_abbr.length, sel_list.length]
        for _per = 1 to per_abbr.length do
            per = per_abbr[_per]
            for _sel = 1 to sel_list.length do
                sel_f = 'AB_Flow_'+sel_list[_sel] + "_" + per
                sel_exp[_per][_sel] = 'if nz('+sel_f+')=0 then null else Format('+sel_f+', "*.")'
            end
        end
        
        //*** Create shading themes for different options ***
        dim orig_themes[per_abbr.length, sel_list.length]
        dim dest_themes[per_abbr.length, sel_list.length]
        
        for _per = 1 to per_abbr.length do
            per = per_abbr[_per]
            for _sel = 1 to sel_list.length do
                Opts = null
                Opts.[Data Source] = "Screen"
                Opts.ThemeName = per + " " + sel_list[_sel] + " Origins"
                Opts.CreateOnly = True
                orig_themes[_per][_sel] = MP.Shading(MP.Views.ZoneData+"."+per+"_"+sel_list[_sel]+"_O", "Optimal", 8, Opts)
                
                Opts.ThemeName = per + " " + sel_list[_sel] + " Destinations"
                dest_themes[_per][_sel] = MP.Shading(MP.Views.ZoneData+"."+per+"_"+sel_list[_sel]+"_D", "Optimal", 8, Opts)
            end
		end
		//Make read only
		CurrentViews = GetViews()
		CurrentViews = CurrentViews[1]
		//if !TrafficMap.Opts.[Input Network] then do
		if True then do //!!! Currently always assume not input network
			for _lyr = 1 to MP.Layers.length do
				lyr = MP.Layers[_lyr][2]
				if ArrayPosition(CurrentViews, {lyr}, ) > 0 then
					SetViewReadOnly(lyr, "True")
			end
			for _vw = 1 to MP.Views.length do
				vw = MP.Views[_vw][2]
				if ArrayPosition(CurrentViews, {vw}, ) > 0 then
					SetViewReadOnly(vw, "True")
			end
		
		end

        SetMapRedraw(MP.Map, "True")
        RunMacro("SetSelectType")
        //If cancelled, we will skip to here
        nomap:
        MP = null //release mapper object 
        
        ShowDbox()
        DisableItem("SelectCreate")
    enditem
    
	
    //StartMethod //Performace Report
    tab "PerformaceReport" prompt: "Report"
    //EndMethod - Performace Report
    
    //StartMethod //Performace Report Create button 
    button "ReportCreate" 2, 8, 30, 2 prompt: "Specify Performance Report >>"
        help: "Create a model summary report" do
		
		//shared ScenName
        
		HideDbox()
        RunMacro("TCB Init")
        //Load performance report object
        SetAlternateInterface(perf_ui)
        InstPerf = null
		InstPerf = CreateObject("Performance", Args) //Args, selected scenario name and scenario directory
        Create_Report = RunDbox("Performance", InstPerf)	//Boolean variable, whether to create the performance report 
		SetAlternateInterface()
		
        if Create_Report then do 					
			ret_value = InstPerf.CreateReport()
			if !ret_value then goto quit
		end	
		
		quit:
        RunMacro("TCB Closing", ret_value, !ret_value)
		ShowDbox()
    enditem	
	
	

	
    //Macro to set the select link map type
    macro "SetSelectType" do
    
        // *** Hide all themes in the map ***
        MP = SelectMap.MP
        MP.ClearThemes(MP.Layers.Links)
        MP.ClearThemes(MP.Layers.Zones)
        
        // *** Theme and label settings for link map ***
        if SelectMap.Type = 1 then do 
            ShowTheme(, link_themes[SelectMap.Opts.Period][sel_val])
            if SelectMap.Opts.Labels then do
                MP.Label(sel_exp[SelectMap.Opts.Period][sel_val], SelLabelOpts)
            end else do
                MP.HideLabels()
            end
        end
        else do 
            ShowTheme(, ft_theme)
        end
        
        //*** Theme and label settings for TAZ O or D map ***
        if SelectMap.Type = 2 or SelectMap.Type = 3 then do
        
            orig_lyr = GetLayer()
            SetLayer(MP.Layers.Zones)
            if SelectMap.Type = 2 then ShowTheme(, orig_themes[SelectMap.Opts.Period][sel_val])
            if SelectMap.Type = 3 then ShowTheme(, dest_themes[SelectMap.Opts.Period][sel_val])
            SetLayer(orig_lyr)
            MP.HideLabels()
        end
        
        //*** Centroid connector and local street display ***
        if SelectMap.Opts.Connectors then do
            SetDisplayStatus(MP.Layers.Links+"|CentroidConnectors", "Active")
        end else do
            SetDisplayStatus(MP.Layers.Links+"|CentroidConnectors", "Invisible")
        end
        
        if SelectMap.Opts.Locals then do
            SetDisplayStatus(MP.Layers.Links+"|Local Streets", "Active")
        end else do
            SetDisplayStatus(MP.Layers.Links+"|Local Streets", "Invisible")
        end
        
        MP.Redraw()
    
    enditem
    
    macro "UpdateSelect" do
        //If we already have a select link map, update it
        if SelectMap.MP != null then do
            RunMacro("SetSelectType")
        end
    enditem
    //EndMethod

    //EndMethod
    
EndDbox

//******************************************************************************
//** Mapper: Flexible map creation                                            **
Class "Mapper"                                                              //** StartClass
// Properties:
// ********************************************************
//
//   -- These properties can be set control map features --
//   -- and should be set prior to calling .Create()     --
//   ------------------------------------------------------
//  .Files.Zones = TAZ layer filename (optional)
//  .MapName = Name of map (defaults to "Map")
//  .Scope = Initial map scope (defaults to network scope)
//  .[Count Field] = Field containing traffic counts for validation
//
//   -- These properties are set as defaults in the init --
//   -- step and tend to vary by model. They can be      --
//   -- overridden after the mapper object is created    --
//   ------------------------------------------------------
//  .Settings.Network.FT
//     .[FT Name] = {int index, int width, string color, string style)
//       - [FT Name] = Descriptive facility type name
//       - index = FT link value
//       - style and color must be available in the Mapper style and color list
//
//   -- These properties are set by methods and should be --
//   -- treated as or read-only                           --
//   -------------------------------------------------------
//  .Files.Network = Network filename
//  .Map = Map handle
//  .Layers = Layers avaialble in the map:
//      .Links = Network link layer
//      .Nodes = Network node layer
//      .Zones = Zone layer
//  .Views = Views avaialble
//      .Flow = Flow view (assignment results)
//      .NetFlow = Network layer joined to the flow view
//      .CompareFlow = Flow view for comparison
//      .NetCompareFlow = Networ+Flow+CompareFlow
//
// Methods
// *********************************************************
//  .Create(base_file, StopRedraw) -> Create a new map
//      base_file = String base layer filename (should be a line layer or
//                  a route system)
//      Boolean StopRedraw = True to supress map redraw, False (default) to 
//                           leave mapredraw on
//  .Connectors(qry, Visible) -> Selects centroid connectors
//      qry = Centroid connector selection set query
//      Visible = True/False: show centroid connectors if True
//  .Select(QryName, qry, Visible, lyr) -> Selects features
        //QryName: Name of the query to create
        //qry: Query for use in hiding links (required)
        //Visible: make the set active (invisible if false)?
        //lyr: layer to act on (defaults to link layer)
//
//  .JoinFlows(FlowFile) -> Join assignment results to the network
//      FlowFile = Assignment result filename
//
//  .JoinCompareFlows(FlowFile) -> Join a "baseline" flow file to the network 
//                                 for comparison to the already joined flow
//                                 file and compute difference expressions.
//      FlowFile = Assignment result filename for comparison
//
//  .Label(expr, Opts) -> Apply labels to the link layer
//      expr = Label expression
//      Opts. (All are optional)
//          .Font = Label Font (See SetLabels() for details)
//          .Color = Label color
//          .CCFont = Font for centroid connector labels
//          .CCColor = Color for centroid connector labels
//
//  .ClearThemes(lyr) -> Clear all themes from the layer
//    string lyr = layer name
//  .Bandwidths(Field, Opts) -> Add a bandwidth theme to the link layer
//      Field = Name of field to control bandwidth theme
//      Opts. (All are optional)
//          .[Data Source] = "Screen" to set sizes based on on-screen values or
//                           "All" (default) to use all values on the network
//          .[Line Style] = Line style to use for theme - only used for an AB/BA 
//                          theme, defaulting to "Double"
//
//  .LOS(Field) --> Add an LOS color theme to the map
//      Field = Field containing LOS values
//
//  .Redraw() -> Redraw the map and update the map toolbar
//
//  .SetDataYear(mdb_file) --> Allow the user to select a data year.  Set the
//                             value in the database file.

    //Initialize:
    // - Check for network existence
    // - Set up default map properties
    init do
    NextStep= "Init"
    
		shared UT
        UT = null
		UT = CreateObject("Utilities")
        self.UT = UT
    
        //Default settings 
        self.MapName = "Map"
        self.Files.Network = null
        //self.Files.Routes = null
        self.Files.Zones = null
        self.Scope = null //defaults to network scope
        self.[Count Field] = null
        
        //Default Network Styles
		//Start of SCAG revisions
        self.Settings.Network.FTField = "AB_Facility_Type"
        self.Settings.Network.CCValue = 100
        self.Settings.Network.FT.Proposed = {null, 0.5, "Red", "Dash"}
        self.Settings.Network.FT.Freeway = {10, 2.5, "Black", "Solid"}
        self.Settings.Network.FT.[Principal Arterial] = {40, 2, "Red", "Solid"}
        self.Settings.Network.FT.[Minor Arterial] = {50, 1.5, "Green", "Solid"}
        self.Settings.Network.FT.[Major Collector] = {60, 1, "Blue", "Solid"}
        self.Settings.Network.FT.[Ramp] = {80, 0, "Black", "Solid"}   
        self.Settings.Network.FT.[Urban Local] = {70, 0, "LtGray", "Solid"}        
        
        self.Settings.Network.FT.[Highway (Outside SOI)] = {11, 2, "LtPink", "Solid"}
        self.Settings.Network.FT.[Arterial (Outside SOI)] = {12, 2, "LtPurple", "Solid"}
        self.Settings.Network.FT.[Rural Arterial (Outside SOI)] = {13, 2, "LtBrown", "Solid"}
        self.Settings.Network.FT.[Local Street] = {60, 0, "LtGray", "Solid"}

        self.Settings.Network.FT.[Centroid Connector] = {100, 0, "Gray", "Dash"}
        
		self.Settings.Network.FT.[Walk Connector] = {49, 0, "Blue", "Dash"}
        self.Settings.Network.FT.[Transit Only] = {999, 0, "LtBlue", "Dash"}
		//End of SCAG revisions


    enditem
    //EndMethod

    //Create a new map
    Macro "Create" (base_file, StopRedraw, InOpts) do
        //Options:
        // - [Close Macro]: Name of a macro to be called on close OR done
    
        if base_file = null or TypeOf(base_file) != 'string' then 
            Throw("Cannot create map: No base layer filename provided")
        if GetFileInfo(base_file) = null then
            Throw(self.UT.StrCombine("Cannot create map. File not found:\n%1%", {base_file}))
        
        t = SplitPath(base_file)
        base_ext = Lower(t[4])
        
        if base_ext != '.rts' and base_ext != '.dbd' then 
            Throw(self.UT.StrCombine("Cannot create map: Base layer must be a " + 
                                "geographic file or route system file.\n%1%", 
                                {base_file}))
    
        //Get route system line layer
        if base_ext = '.rts' then do
            self.Files.Routes = base_file
            t = GetRouteSystemInfo(base_file)
            self.Files.Network = t[1]
            t = t[3] //info Opts
            self.Layers.Routes = t.Name  //Rts layer name in file
            t = null
        end
        else
            self.Files.Network = base_file
    
        //Use default scope if not defined externally
        if self.Scope = null then do
            t = GetDBInfo(self.Files.Network)
            self.Scope = t[1]
            t = null
        end
        
        //Create map
        Opts = null
        Opts.Scope = self.Scope
        CM = InOpts.[Close macro]
        if CM != null then do
            Opts.[Close macro] = CM
            Opts.[Done macro] = CM
        end
        self.Map = CreateMap(self.MapName, Opts)
        if StopRedraw then SetMapRedraw(self.Map, "False")
        
        //Zones
        if self.Files.Zones != null and 
           GetFileInfo(self.Files.Zones) != null then do
            {zone_lyr} = GetDBLayers(self.Files.Zones)
            self.Layers.Zones = AddLayer(self.Map, zone_lyr, 
                                         self.Files.Zones, zone_lyr)
            RunMacro("G30 new layer default settings", self.Layers.Zones)
            taz_color = self.Colors("LtOrange")
			SetLayer(self.Layers.Zones)
			SelectNone("Selection")
            SetLineColor(self.Layers.Zones+"|", taz_color)
            SetLineWidth(self.Layers.Zones+"|", 3.5)
        end
        
        //Routes
        if base_ext = '.rts' then do
            lyrs = AddRouteSystemLayer(self.Map, self.Layers.Routes, 
                                       self.Files.Routes,)
            RunMacro("Set Default RS Style", lyrs, "TRUE", "TRUE")
            self.Layers.Routes = lyrs[1]
            self.Layers.Stops  = lyrs[2]
            //Not using physical stops
            self.Layers.Nodes = lyrs[4]
            self.Layers.Links = lyrs[5]
            lyrs = null
			
			SetLayer(self.Layers.Nodes)
			SelectNone("Selection")
			SetLayer(self.Layers.Links)
			SelectNone("Selection")
        end
        //Network
        else do
            {node_lyr, link_lyr} = GetDBLayers(self.Files.Network)
            self.Layers.Nodes = AddLayer(self.Map, node_lyr, 
                                         self.Files.Network, node_lyr)
            RunMacro("G30 new layer default settings", self.Layers.Nodes)
            SetLayerVisibility(self.Map+"|"+self.Layers.Nodes, "False")
            self.Layers.Links = AddLayer(self.Map, link_lyr, 
                                         self.Files.Network, link_lyr)
            RunMacro("G30 new layer default settings", self.Layers.Links)
			SetLayer(self.Layers.Nodes)
			SelectNone("Selection")
			SetLayer(self.Layers.Links)
			SelectNone("Selection")
        end

    enditem
    //EndMethod
    
    //Create a centroid connector selection set
    Macro "Connectors" (qry, Visible) do
	
        SetLayer(self.Layers.Links)
        CentroidCount = SelectByQuery("CentroidConnectors", "Several", qry, )
        if CentroidCount > 0 then do
            if Visible then 
                SetDisplayStatus(self.Layers.Links+"|CentroidConnectors", "Active")
            else
                SetDisplayStatus(self.Layers.Links+"|CentroidConnectors", "Invisible")
        end
    
    enditem
    //EndMethod
    
    //Create a set, make visible or invisible
    Macro "Select" (QryName, qry, Visible, InOPts) do
        //QryName: Name of the query to create
        //qry: Query for use in hiding links (required)
        //Visible: make the set active (invisible if false)?
        // Options:
        //   Layer: layer to act on (defaults to link layer)
        //   LineType = {style, color, width}
        
        if InOPts.Layer != null then lyr = InOPts.Layer
        else lyr = self.Layers.Links
	
        orig_lyr = GetLayer()
        SetLayer(lyr)
        cnt = SelectByQuery(QryName, "Several", qry, )
        
        //Quit if nothing selected
        if nz(cnt) = 0 then Return()

        //Set visibility
        if Visible then 
            SetDisplayStatus(self.Layers.Links+"|"+QryName, "Active")
        else
            SetDisplayStatus(self.Layers.Links+"|"+QryName, "Invisible")
        
        //Set styles if enabled
        LT = InOPts.LineType
        if LT != null and TypeOf(LT) = 'array' and LT.length = 3 then do
        
            SetLineStyle(lyr+"|"+QryName, LT[1])
            SetLineColor(lyr+"|"+QryName, LT[2])
            SetLineWidth(lyr+"|"+QryName, LT[3])
        
        end
        
        SetLayer(orig_lyr)
    
    enditem
    //EndMethod

    //Join assignment results to the network
    Macro "JoinFlows" (FlowFile) do
    
        if GetFileInfo(FlowFile) != null then do
            self.Views.Flow = OpenTable("Flow", "FFB", {FlowFile, })
            self.Views.NetFlow = JoinViews("Network+Flow", 
                                           self.Layers.Links+".ID", 
                                           self.Views.Flow+".ID1", )
            //Close the flow view so it doesn't remain after the map is closed
            CloseView(self.Views.Flow)
        end
        else do
            if TypeOf(FlowFile) = TypeOf("string") then
                Throw("Cannot join flow file to map - Flow file not Found\n"+FlowFile)
            else 
                Throw("Cannot join flow file to map - Incorect function argument FlowFile")
        end
        
        Return(self.Views.NetFlow)
        
    enditem
    //EndMethod
	
    //Join data to the TAZ Layer
    Macro "JoinData" (DataFile, joinID) do
    //Data file can be:
    // - DBASE / FFB / FFA / CSV
    // - Access in the form "*.dbd|TableName"
    
        //Verify that the zone layer is in the map
        if self.Layers.Zones = null then do
            ShowMessage("Cannot join data - no zone layer in map")
            Return()
        end
        
        //Identify data fileype
        fspec = SplitString(DataFile)
        ftype = RunMacro("G30 table type", fspec[1])
        
        //Access needs to know which unique ID to use
        if ftype = "ACCESS" then fspec = fspec + {joinID}
        
        //Open the table for joining
        self.Views.Data = OpenTable("Data", ftype, fspec, )
        if self.Views.Data = null then do
            ShowMessage("Error joining TAZ data to zone layer.")
            Return()
        end
        
        //Join 
        self.Views.ZoneData = JoinViews("Zones+Data", self.Layers.Zones+".TAZ", 
                                        self.Views.Data+"."+joinID, )
    
    enditem
    //EndMethod
    
    //Join all select link flows into a temporary memory view
    Macro "JoinSelect" (flow_files, per_list, sel_list)do
    
        SetVs = null
        FldSpec = {{"ID1", "Integer", 10, 0, "True"}}
        for _per = 1 to per_list.length do
            per = per_list[_per]
            join_vw = self.JoinFlows(flow_files[_per])
            if _per = 1 then do
                SetVs.ID1 = GetDataVector(join_vw+"|", "ID", )
            end
            
            for sel in sel_list do
                flds = {'AB_Flow_'+sel, 'BA_Flow_'+sel}
                Vs = GetDataVectors(join_vw+"|", flds, )
                SetVs.(flds[1]+"_"+per) = Vs[1]
                SetVs.(flds[2]+"_"+per) = Vs[2]
                
                FldSpec = FldSpec + {{flds[1]+"_"+per, "Real", 10, 2}, 
                                     {flds[2]+"_"+per, "Real", 10, 2}}
                
            end
            
            //Close the joined view, unless is is in use by another map
            on error goto NoCloseJoin
            CloseView(join_vw)
            NoCloseJoin:
            self.Views.NetFlow = null
            
        end
        
        sel_mem = CreateTable("SelectFlowsMem", , "MEM", FldSpec)
        AddRecords(sel_mem, , , {{"Empty Records", SetVs.ID1.length}})
        SetDataVectors(sel_mem+"|", SetVs, )
        self.Views.SelectFlows = JoinViews("SelectFlows", self.Layers.Links+".ID", sel_mem+".ID1", )
        CloseView(sel_mem)
        
        Return(MP.Views.SelectFlows)
    
    enditem //EndMethod
    
    //Get a select link query settings (separate dialog box)
    Macro "SelectMapSettings" (qry_file) do
        Return(RunDbox("SelectMapSettings", qry_file))
    EndItem
    //EndMethod
    
    //Get a scenario for comparison
    Macro "GetScenario" do
        Return(RunDbox("SelectCompare"))
    EndItem
    //EndMethod
	
	//Ask the user to select a data year, set database to this value
	Macro "SetDataYear" (mdb_file, avail_tname, act_tname) do
		Return(RunDbox("SetDataYear", mdb_file, avail_tname, act_tname))
	EndItem
	//EndMethod
    
    //Join assignment results to the network for comparison and setup map
    Macro "CompareFlows" (FlowFile, threshold) do
    
        //Default threshold for no difference
        if threshold = null then threshold = 500
        threshold = String(threshold)
    
        if GetFileInfo(FlowFile) = null or
           self.Views.Flow = null then 
            Throw("Cannot join flow file for comparison")
        
        
        //Open and join the comparison flow
        self.Views.CompareFlow = OpenTable("CompareFlow", "FFB", {FlowFile, })
        self.Views.NetCompareFlow = JoinViews("Network+Flow+CompareFlow", 
                                               self.Views.NetFlow+".ID", 
                                               self.Views.CompareFlow+".ID1", )
        CloseView(self.Views.CompareFlow)
        
        //Compute expressions
        f_vw = self.Views.Flow
        c_vw = self.Views.CompareFlow
        join_vw = self.Views.NetCompareFlow
        diff_expr = "nz(" + f_vw + ".TOT_Flow) - nz(" + 
                            c_vw + ".TOT_Flow)"
        abs_expr = "ABS(DIFF)"
        color_expr = "if ABSDIFF < " + threshold + " then 0 else DIFF/ABSDIFF"
        CreateExpression(join_vw, "DIFF", diff_expr, )
        CreateExpression(join_vw, "ABSDIFF", abs_expr, )
        CreateExpression(join_vw, "DIFFCOLOR", color_expr, )
        
        //Create color theme
        whr_colors = {self.Colors("Gray"),  //Other
                      self.Colors("Blue"),  //Down (-1)
                      self.Colors("Gray"),  //No Change (0)
                      self.Colors("DkRed")} //Up (1)
                      
        whr_names = {"Minmial Change",
                     "Decrease in Volume", 
                     "Minmial Change",
                     "Increase in Volume"}


        whr_th = CreateTheme("Direction of Change", join_vw+".DIFFCOLOR", 
                             "Categories", whr_colors.length - 1, )
        SetThemeLineColors(whr_th, whr_colors)
        ShowTheme( , whr_th)
        
        //Legend settings (Color)
		class_cnt = GetThemeClassLabels(whr_th)
		if class_cnt.length < 4 then do
			ShowMessage("Warning - insufficient differences between themes to create complete map.  Use results carefully!")
		end
		else do
			SetLegendDisplayStatus(whr_th+"|3", "False")
		end
        SetThemeClassLabels(whr_th, whr_names)
        
        //Scaled Symbol Theme
        self.Bandwidths(join_vw+".ABSDIFF")
    
    EndItem
    //EndMethod

    //Apply labels to the link layer
    Macro "Label" (expr, InOpts) do
    //  Options:
    // - Suffix: Suffix to add to the label name (defaults to a random number)
    // - ExpressionView: View to use for expression creation (e.g., joined vw)
    // - CC.Expression: Expression to use for centroid connectors
    // ** Below also have CC.* variantes **
    // - [Priority Expression]: Priority expression
    // - Font: Label font (defaults to "Arial|8.5", "Arial |7" for CC)
    // - Color: Label color (defaults to black, Gray for CC)
    
    
        SetView(self.Layers.Links)
        //!!! Two Way: limited functionality.  
        // - Cannot use a different CC expression
        
        //assign a unique label name
        lbl_suffix = InOpts.Suffix
        if lbl_suffix = null then do
            r = R2I(Round(RandomNumber() * 100, 0))
            lbl_suffix = Format(r, "*.")
        end
        
        
        //Check for expression view override
        if InOpts.ExpressionView != null then
            vw = InOpts.ExpressionView
        else
            vw = self.Views.NetFlow
            
        //Check for AB/BA Labels
        two_way = False
        if Position(expr, "AB") > 0 then do
        
            BAexpr = Substitute(expr, "AB", "BA", )
            on Error goto NoTwoWay
            VerifyExpression(vw, BAexpr)
            two_way = True
        end
        NoTwoWay:
        on Error default
        
        //Destroy old expressions if current label name is already present
        all_exp = GetExpressions(vw)
        if two_way then do
            AB_tmp = "AB_lbl_"+lbl_suffix
            if ArrayPosition(all_exp, {AB_tmp}, ) > 0 then DestroyExpression(vw+"."+AB_tmp)
            BA_tmp = "BA_lbl_"+lbl_suffix
            if ArrayPosition(all_exp, {BA_tmp}, ) > 0 then DestroyExpression(vw+"."+BA_tmp)
        end else do
            tmp = "lbl_"+lbl_suffix
            if ArrayPosition(all_exp, {tmp}, ) > 0 then DestroyExpression(vw+"."+tmp)
        end
        
        //Always destroy an old priority expression
        if ArrayPosition(all_exp, {"pri"}, ) > 0 then DestroyExpression(vw+".pri")
        
        //Set up label expressions
        if two_way then do
            lbl = CreateExpression(vw, "AB_lbl_"+lbl_suffix, expr, )
            BAlbl = CreateExpression(vw, "BA_lbl_"+lbl_suffix, BAexpr, )
            cc_lbl = lbl //!!! todo: one-way doesn't allow for a different centroid connector expression
        end else do
            lbl = CreateExpression(vw, "lbl_"+lbl_suffix, expr, )
            if InOpts.CC.Expression != null then 
                cc_lbl = CreateExpression(vw, "cc_lbl_"+lbl_suffix, InOpts.CC.Expression, )
            else
                cc_lbl = lbl
        end
        
        //Set up priority expressions
        ft_fld = self.Fields.FTField
        if InOpts.[Priority Expression] != null then do
			pri_exp = InOpts.[Priority Expression]
            pri = CreateExpression(vw, "pri", pri_exp, )
		end
		else pri = null
        
        if InOpts.CC.[Priority Expression] != null then do
			pri_exp = InOpts.CC.[Priority Expression]
            CCpri = CreateExpression(vw, "CCpri", pri_exp, )
		end
        else CCpri = pri

        //Load label options
        Opts = null
        Opts.[Priority Expression] = "-"+pri
        if InOpts.Font != null then 
            Opts.Font = InOpts.Font
        else
            Opts.Font = "Arial|8.5"
        if InOpts.Color != null then
            Opts.Color = InOpts.Color
        else
            Opts.Color = self.Colors("Black")
            
        CCOpts = null
        CCOpts.[Priority Expression] = "-"+CCpri
        if InOpts.CC.Font != null then 
            CCOpts.Font = InOpts.CC.Font
        else
            CCOpts.Font = "Arial|7"
        if InOpts.CC.Color != null then
            CCOpts.Color = InOpts.CC.Color
        else
            CCOpts.Color = self.Colors("Gray")

        Opts.Rotation = "True"
        CCOpts.Rotation = "True"
        Opts.Visibility = "On"
        CCOpts.Visibility = "On"
        Opts.[Set Priority] = 5
        CCOpts.[Set Priority] = 7
        
        if two_way then do
            Opts.[Left/Right] = "True"
            CCOpts.[Left/Right] = "True"
        end
        
        //Activate Labels
        SetLabels(self.Layers.Links+"|", lbl, Opts)
        
        //Different settings for centroids (if selected)
        cnt = 0
        on NotFound goto next1
        cnt = GetSetCount("CentroidConnectors")
        next1:
        on NotFound default
        if cnt > 0 then 
            SetLabels(self.Layers.Links+"|CentroidConnectors", cc_lbl, CCOpts)
        
        //Highlight for validation
        if InOpts.Highlight then do
            SetView(self.Layers.Links)
			/*
            SelectByQuery("OK", "Several", 
                          "Select * Where abs(TOT_Flow- "+self.[Count Field]+")<3000 or " + 
                          "((TOT_Flow/"+self.[Count Field]+")>=0.9 and " + 
                          "(TOT_Flow/"+self.[Count Field]+"<=1.1))", )
            SelectByQuery("High", "Several", 
                          "Select * Where (TOT_Flow- "+self.[Count Field]+")> 3000 and " + 
                          "(TOT_Flow/"+self.[Count Field]+">1.1)", )
            SelectByQuery("Low", "Several", 
                          "Select * Where (TOT_Flow- "+self.[Count Field]+")< -3000 and " + 
                          "(TOT_Flow/"+self.[Count Field]+"<0.9)", )
						  */
						  
            SelectByQuery("OK", "Several", 
                          "Select * Where abs(TOT_Flow- "+self.[Count Field]+")<3000 or " + 
                          "((TOT_Flow/"+self.[Count Field]+")>=0.8 and " + 
                          "(TOT_Flow/"+self.[Count Field]+"<=1.2))", )
            SelectByQuery("High", "Several", 
                          "Select * Where (TOT_Flow- "+self.[Count Field]+")> 3000 and " + 
                          "(TOT_Flow/"+self.[Count Field]+">1.2)", )
            SelectByQuery("Low", "Several", 
                          "Select * Where (TOT_Flow- "+self.[Count Field]+")< -3000 and " + 
                          "(TOT_Flow/"+self.[Count Field]+"<0.8)", )
                          
            //Set up defaults for new selection sets
            SetDisplayStatus("OK", "Active")
            SetDisplayStatus("Low", "Active")
            SetDisplayStatus("High", "Active")
            
            SetLineStyle(self.Layers.Links+"|OK", null)
            SetLineStyle(self.Layers.Links+"|Low", null)
            SetLineStyle(self.Layers.Links+"|High", null)
            SetLineColor(self.Layers.Links+"|OK", null)
            SetLineColor(self.Layers.Links+"|Low", null)
            SetLineColor(self.Layers.Links+"|High", null)
            
            //Set up labels for highlight selection sets
            fill_sty = RunMacro("G30 setup fill styles")
            LabelOptsColor = CopyArray(Opts)
            LabelOptsColor.[Frame Border Style] = self.Styles("None")
            LabelOptsColor.[Frame Border Color] = ColorRGB(0, 0, 0)
            LabelOptsColor.[Frame Border Width] = 0
            LabelOptsColor.[Frame Fill Color] = ColorRGB(65535, 65535, 0)
            LabelOptsColor.[Frame Fill Style] = fill_sty[2]
            LabelOptsColor.[Frame Type] = "rounded rectancle"
            LabelOptsColor.Framed = "True"
            SetLabels(self.Layers.Links+"|OK", lbl, LabelOptsColor)
            LabelOptsColor.[Frame Fill Color] = ColorRGB(47360, 56320, 65280)
            SetLabels(self.Layers.Links+"|Low", lbl, LabelOptsColor)
            LabelOptsColor.[Frame Fill Color] = ColorRGB(65280, 47360, 47360)
            SetLabels(self.Layers.Links+"|High", lbl, LabelOptsColor)
            LabelOptsColor = null
        end //Highlight
        
        self.LabelExpression = lbl
        
    enditem
    //EndMethod
    
    //Hide all layers for all sets on the identified layer
    Macro "HideLabels" (lyr) do
    // lyr: Layer to operate on.  Defaults to link layer
    
        if lyr = null then lyr = self.Layers.Links
        sets = {null} + GetSets(lyr)
        
        for set in sets do
            {vis} = GetLabelOptions(lyr+"|"+set, {"Visibility"})
            if vis then do
                SetLabelOptions(lyr+"|"+set, {{"Visibility", "Off"}})
            end
        end
    enditem //EndMethod
    
    //Clear all themes from the link layer
    Macro "ClearThemes" (lyr) do
        
        //Hide visible themes
        orig_lyr = GetLayer()
        SetLayer(lyr)
        th = GetDisplayedThemes(self.Map+"|"+lyr+"|")
        for i = 1 to th.Length do
            HideTheme(null, th[i])
        end
        
        //Don't destroy the themes - they may be used by another map in the 
        //workspace
        
        SetLayer(orig_lyr)
    
    enditem
    //EndMethod
   
    //Clear all selection sets from the specified layer
    Macro "ClearSets" (lyr) do
        orig_lyr = GetLayer()
        SetLayer(lyr)
        sets = GetSets(lyr)
        for i = 1 to sets.length do
            if sets[i] = "Selection" then
                SelectNone("Selection")
            else 
                DeleteSet(sets[i])
        end
        SetLayer(orig_lyr)
    EndItem
    //EndMethod
    
    //Add a bandwidth theme to the link layer
    Macro "Bandwidths" (FieldSpec, InOpts) do
        //InOpts:
        // - Data Source: "Screen" to override default "All"
        // - Line Style: Line style to override devault "Double"
        // - ThemeName: Name to override default "Volume Bandwidths"
        // - CreateOnly: True to create but not show the theme
    
        Opts = null
        if InOpts.[Data Source] = "Screen" then 
            Opts.[Data Source] = "Screen"
        if InOpts.[Line Style] = null then 
            line_style = "Double"
        else
            line_style = InOpts.[Line Style]
        if InOpts.ThemeName != null and TypeOf(InOpts.ThemeName) = TypeOf('str') then 
            ThemeName = InOpts.ThemeName
        else 
            ThemeName = "Volume Bandwidths"
            
        bd_theme = CreateContinuousTheme(ThemeName, {FieldSpec}, Opts)
        ShowTheme( , bd_theme)
        SetLegendDisplayStatus(bd_theme, "False")
        
        //Set the line style if the field starts with AB or BA
        t = ParseString(FieldSpec, '.')
        t = Upper(Left(t[2], 2))
        if t = 'AB' or t = 'BA' then do
            //SetThemeLineStyles(bd_theme, {self.Styles('Double')})
            SetThemeLineStyles(bd_theme, {self.Styles(line_style)})
            SetThemeLineColors(bd_theme, {self.Colors('BrtGreen')})
            SetThemeLineWidths(bd_theme, {3})
        end
        
        if !InOpts.CreateOnly then do
            ShowTheme( , bd_theme)
        end
        
        Return(bd_theme)
    
    enditem
    //EndMethod

    //Add a color theme to the zone layer
    Macro "Shading" (FieldSpec, method, classes , InOpts) do
        //InOpts:
        // - Data Source: "Screen" to override default "All"
        // - Color: Color to override default blue
        // - ThemeName: Name to override default "Density"
        // - CreateOnly: True to create but not show the theme
        
        ZoneLayer = self.Layers.Zones
        orig_lyr = GetLayer()
        SetLayer(ZoneLayer)
        Opts = null
        if InOpts.[Data Source] = "Screen" then 
            Opts.[Data Source] = "Screen"
        if InOpts.[Color] = null then Color = self.Colors("Blue")
        else Color = InOpts.[Color]
        if InOpts.ThemeName != null and TypeOf(InOpts.ThemeName) = TypeOf('str') then 
            ThemeName = InOpts.ThemeName
        else 
            ThemeName = "Density"
            
        sh_theme = CreateTheme(ThemeName, FieldSpec, method, classes, Opts)
        SetLegendDisplayStatus(sh_theme, "True")

        //Setup color gradient
        palette = GeneratePalette(self.Colors("White"), Color, classes-1, )
        for cls = 1 to palette.length do
            cls_id = ZoneLayer + "|" + sh_theme + "|" + String(cls)
            SetFillColor(cls_id, palette[cls])
            SetFillStyle(cls_id, solid)
        end
        
        if !InOpts.CreateOnly then do
            ShowTheme( , sh_theme)
        end
        
        SetLayer(orig_lyr)
        Return(sh_theme)
        
    enditem
    //EndMethod
    
    //Add an LOS color theme to the map
    Macro "LOS" (Field) do
    
        //Setup colors and names
        los_colors =   {self.Colors("Gray"),     //Other
                        self.Colors("Green"),    //A
                        self.Colors("Green"),    //B
                        self.Colors("Green"),    //C
                        self.Colors("Orange"),   //D
                        self.Colors("Orange"),   //E
                        self.Colors("Red"),      //F
                        self.Colors("Gray")}     //n/a
                        
        los_names = {"Not Computed",
                     "Uncongested (A - C)",
                     "Uncongested (A - C)",
                     "Uncongested (A - C)",
                     "Congesting (D - E)",
                     "Congesting (D - E)",
                     "Congested (F)", 
                     "Not Computed"}
        
        //Create and show the theme
        los_th = CreateTheme("LOS", self.Views.NetFlow+"."+Field, 
                             "Categories", los_names.length -1, ) 
                             //-1 because names also include "Other"
                             
        SetThemeLineColors(los_th, los_colors)
        ShowTheme( , los_th)
        
        //Hide redundant legend entries
        SetLegendDisplayStatus(los_th+"|1", "False")
        SetLegendDisplayStatus(los_th+"|2", "False")
        SetLegendDisplayStatus(los_th+"|3", "False")
        SetLegendDisplayStatus(los_th+"|5", "False")
        
        SetThemeClassLabels(los_th, los_names)   
    
    enditem
    //EndMethod
    
    //Add a V/C ratio color theme to the map
    Macro "VOC" (Field) do
    
        if Field = null then Field = "AB_VOC"
    
        //LOS VC values based on arterial cutpoints
    /* !!!
        vc_values = {{0.00, "True", 0.51, "False"},   //A
                     {0.51, "True", 0.67, "False"},   //B
                     {0.67, "True", 0.79, "False"},   //C
                     {0.79, "True", 0.90, "False"},   //D
                     {0.90, "True", 1.00, "False"},   //E
                     {1.00, "True", 9999, "False"}}   //F

        vc_colors = {self.Colors("Gray"),       //Other
                     self.Colors("LtBlue"),     //A
                     self.Colors("BlueGreen"),  //B
                     self.Colors("Green"),      //C
                     self.Colors("Orange"),     //D
                     self.Colors("DkOrange"),   //E
                     self.Colors("Red")}        //F
          */
          
        vc_values = {{0.00, "True", 0.51, "False"},   //A
                     {0.51, "True", 0.67, "False"},   //B
                     {0.67, "True", 0.79, "False"},   //C
                     {0.79, "True", 0.90, "False"},   //D
                     {0.90, "True", 1.00, "False"},   //E
                     {1.00, "True", 9999, "False"}}   //F

        vc_colors = {self.Colors("Gray"),       //Other
                     self.Colors("Green"),      //A
                     self.Colors("Green"),      //B
                     self.Colors("Green"),      //C
                     self.Colors("Orange"),     //D
                     self.Colors("DkOrange"),   //E
                     self.Colors("Red")}        //F


          
        //Create and show the theme
        vc_th = CreateTheme("V/C", self.Views.NetFlow+"."+Field, 
                            "Manual", vc_values.length, {{"Values", vc_values}})
                            
        SetThemeLineColors(vc_th, vc_colors)
        ShowTheme( , vc_th)
                            
    
    EndItem
    //EndMethod

    //Redraw the map and update the map toolbar
    Macro "Redraw" do
        RedrawMap(self.Map)
        RunMacro("G30 update map toolbar")
    enditem
    //EndMethod

	//**************************************************************************
	//** Macro to apply a FT color theme and save settings to the DBD file
	Macro "NetworkSetting" (dbd_file, close, save) do
	//** String dbd_file = Name of the network file
    //** Boolean close = True to close the map on finish, or False (default) to 
    //**  leave the map open
    //** Boolean save = True to save settings to the "stg" file, or False to 
    //** simply apply the settings
    //**
    //** This macro relies on the settings specified in the Settings property
    //**************************************************************************
	
NextStep= "Map Setup"
        //Settings
        NetSet = self.Settings.Network
    
		//Open the dbd file in a map
		self.Create(dbd_file, close) //if close=True, then StopRedraw parameter is True
		
		//Remove any pre-existing link themes
		self.ClearThemes(self.Layers.Links)
        self.ClearThemes(self.Layers.Nodes)
        
//EndStep
NextStep= "FT Theme"

        self.FTTheme()
        
//EndStep
NextStep= "Link Selection Sets"

        //Clear/remove all sets
        SetLayer(self.Layers.Links)
        self.ClearSets(self.Layers.Links)
        
        //Centroid connectors
        cc_qry = "Select * Where " + NetSet.FTField + " = " + 
                 String(NetSet.CCValue)
        SelectByQuery("CentroidConnectors", "Several" ,cc_qry, )
        SetDisplayStatus(self.Layers.Links+"|CentroidConnectors", "Active")
        
//EndStep
NextStep= "Node Theme"     

        //Set nodes with ZONE > 0 to a blue triangle, other nodes to an orange
        //  target.  Leaves the basic node style unchanged.
        SetLayerVisibility(self.Layers.Nodes, "True")
        SetLayer(self.Layers.Nodes)
        
        Opts = null
        Opts.Values = {{0, "True", 999999999999999999, "True"}}
        
        node_th = CreateTheme("Node Type", self.Layers.Nodes+".ZONE", "Manual", 1, Opts)
        
        SetThemeClassLabels(node_th, {"Nodes", "Centroids"})
        SetThemeIcons(node_th, {{"Font Character", "Caliper Cartographic|8", 38},
                                {"Font Character", "Caliper Cartographic|6", 39}})
        SetThemeIconColors(node_th, {self.Colors("Orange"), self.Colors("Blue")})
        ShowTheme( , node_th)
    
//EndStep
NextStep= "Node Selection Sets"  

        self.ClearSets(self.Layers.Nodes)
        SetLayer(self.Layers.Nodes)
        c_qry = "Select * Where ZONE > 0"
        SelectByQuery("Centroids", "Several" ,c_qry, )
        SetDisplayStatus(self.Layers.Nodes+"|Centroids", "Active")

//EndStep
NextStep= "Save Settings"

		//Save the theme as the default display for the link layer
        self.Redraw()
		t = SplitPath(dbd_file)
		sty_file = t[1]+t[2]+t[3]+".sty"
		SetDefaultDisplay(self.Layers.Links, sty_file)
        sty_file = t[1]+t[2]+t[3]+"_.sty"
		SetDefaultDisplay(self.Layers.Nodes, sty_file)
		
//EndStep
NextStep= "Close the map"
        
        if close then CloseMap(self.Map)
		
//EndStep	
	//EndMethod
    EndItem
    
	//**************************************************************************
	//** Macro to apply a FT color theme to an open map
	Macro "FTTheme" (InOpts) do
        //InOpts:
        // - CreateOnly: True to create but not show the theme
    
        NetSet = self.Settings.Network
    
        SetLayer(self.Layers.Links)
		//Create the link theme (Include all values up to 99 unique numbers)
		ft_theme = CreateTheme("Roadways", self.Layers.Links+"."+NetSet.FTField, 
                               "Categories", 99, )
		ft_thvals = GetThemeClassValues(ft_theme) //First element null for other
        
        //Re-format FT options for use in theme settings.  While we're at it, 
        //  make sure that the passed values line up correctly with the actual
        //  theme values.  This will help prevent problems when a particular 
        //  value is missing or when there is an extra value in the file.
        dim ft_labels[ft_thvals.Length]
        dim ft_widths[ft_thvals.Length]
        dim ft_colors[ft_thvals.Length]
        dim ft_styles[ft_thvals.Length]
        Ks = self.UT.Keys(NetSet.FT)
        Vs = self.UT.Values(NetSet.FT)
        for i = 1 to Ks.Length do
            k = Ks[i] //FT Description
            v = Vs[i]
            //v = {FTval, width, color, style}
            if v[1] != null then idx = ArrayPosition(ft_thvals, {v[1]}, )
            else idx = 1 //Other in the first slot
            if idx > 0 then do
                ft_labels[idx] = string(v[1]) + " - " + k
                ft_widths[idx] = v[2]
                ft_colors[idx] = self.Colors(v[3])
                ft_styles[idx] = self.Styles(v[4])
            end
        end
        
        //Add labels for FT values not included in the options
        for i = 1 to ft_labels.length do
            if ft_labels[i] = null then do
                if i = 1 then ft_labels[i] = "Other"
                else ft_labels[i] = String(ft_thvals[i])
            end
        end
		
        //Apply settings
		SetThemeClassLabels(ft_theme, ft_labels)
		SetThemeLineStyles(ft_theme, ft_styles)
        SetThemeLineColors(ft_theme, ft_colors)
		SetThemeLineWidths(ft_theme, ft_widths)
		
		//Show the theme
        if !InOpts.CreateOnly then do
            ShowTheme( , ft_theme)
        end
        
        Return(ft_theme)
        
    enditem
    //EndMethod
    
    
    //Macro to return a color based on a name (or a list of colors)
    Macro "Colors" (color) do
        //Define a colors attribute for easy reference
        Colors = null
        Colors.White = ColorRGB(65535, 65535, 65535)
        Colors.Gray = ColorRGB(32768,32768,32768) //The standard "TransCAD Gray"
        Colors.LtGray = ColorRGB(49152,49152,49152)
        Colors.Black = ColorRGB(0, 0, 0)
        Colors.BlueGreen = ColorRGB(0, 49152, 49152)
        Colors.Red = ColorRGB(65535, 0, 0)
        Colors.DkRed = ColorRGB(49152, 1228, 0)
        Colors.BrtGreen = ColorRGB(0, 65535, 0)
        Colors.DkGreen = ColorRGB(0, 18384, 0)
        Colors.Green = ColorRGB(0, 49512, 0)
        Colors.Blue = ColorRGB(0, 0, 65535)
        Colors.LtBlue = ColorRGB(0, 44032, 65535)
        Colors.DkBlue = ColorRGB(0, 0, 32768)
        Colors.Orange = ColorRGB(65535, 49152, 0)
        Colors.LtOrange = ColorRGB(65280, 59648, 42496)
        Colors.DkOrange = ColorRGB(65535, 32768, 0)
		Colors.Cyan = ColorRGB(0, 65535, 65535)
        Colors.LtPink = ColorRGB(65280, 40704, 65280)
        Colors.LtPurple = ColorRGB(42496, 42496, 65280)
        Colors.LtBrown = ColorRGB(56832, 48384, 50176)
        
        if color = null or Colors.(color) = null then
            Return(Colors)
        else 
            Return(Colors.(color))
    EndItem
    //EndMethod
    
    //Macro to return a style based on a name (or a list of styles)
    //Pass a style name to get a style name, or pass "List" to get a list of
    //available styles.  Returns null if an invalid style is passed.
    Macro "Styles" (style) do
        sty = RunMacro("G30 setup line styles")
        Styles.None = sty[1]
        Styles.Solid = sty[2]
        Styles.Dot = sty[3]
        Styles.Dash = sty[6]
        Styles.Double = sty[69]
        
        if Lower(style) = "List" then
            Return(Styles)
        if style = null then
            Return(null)
        else 
            Return(Styles.(style))
    enditem
    //EndMethod
        
EndClass

//******************************************************************************
//** Mapper dialog box: ask for a scenario for comparison                     **
Dbox "SelectCompare" title: "Comparison Map"                                //**
// NOTE - this access the scenario files through a shared variable, AND       **
//        references file keys by name -                                      **
// THIS IS A "HIGH MAINTENENCE" DIALOG BOX                                    **
//******************************************************************************

    init do
        //Get completed scenarios
        shared ScenArr, Args, scenario_list, DashScen
        shared ScenFlag
        
        ScenList = null
        FlowList = null
        
        //Get selected scenario flow file for reference
        //Args = ScenArr[ScenFlag[1]][2]
		ScenName = scenario_list[DashScen]
        //scen_flow = Args.Output.PST.DailyFlow.Value
		scen_flow = Args.[Hwy Day Final Flow Table]
        
        for i = 1 to ScenArr.Length do
        
        //Don't add current scenario
            if i <> ScenFlag[1] then do
                //Don't add if daily flow file is not present
                //Args = ScenArr[i][2]
				Args_base = RunMacro("TCP Convert to Argument Options", ScenArr[i][5])
                base_name = ScenArr[i][1]
                //base_flow = Args.Output.PST.DailyFlow.Value
				base_flow = Args_base.[Hwy Day Final Flow Table]
                
                if GetFileInfo(base_flow) <> null then do
                    ScenList = ScenList + {base_name}
                    FlowList = FlowList + {base_flow}
                end
            end
        end //end loop over available scenarios
        
        if ScenList = null then do
            ShowMessage("No completed scenarios exist for comparison")
            Return()
        end
        
        BaseFlag = 1  //Start with first scenario in the list
    enditem //init
    
    text 1, 1 Variable: "Select a baseline scenario for comparison:"
    popdown Menu "Baseline" 3, 2.5, 30 List: ScenList Variable: BaseFlag
    button "OK" 15, 5.5, 10, 1.5 do
        if base_flow = scen_flow then do
            ShowMessage("Scenarios reference the same output flow file.  Cannot create comparison map.")
            Return()
        end
        Return(FlowList[BaseFlag])
    enditem
    button "Cancel" 26, 5.5, 10, 1.5 do
        Return()
    enditem

EndDbox


//******************************************************************************
//** Mapper dialog box: select link query map settings                        **
Dbox "SelectMapSettings" (qry_file)                                         //**
     title: "Select Link/Node Map"                                          //**

    init do
        //Get list of querys
        shared UT
        sel_list = UT.SelectList(qry_file)
        sel_val = 1
    
    enditem
    
    text "Query: " 1, 1
    popdown menu "QueryName" 13, 1, 15, 5 list: sel_list variable: sel_val
    
    button "OK" 16, 5, 7, 1.5 do
        RetOpts = null
        RetOpts.QueryName = sel_list[sel_val]
        Return(RetOpts)
    enditem
    button "Cancel" after, same, 7, 1.5 do
        Return()
    enditem
    

EndDbox

//******************************************************************************
//** Mapper dialog box: select data year (for TAZ data edit)                  **
Dbox "SetDataYear" (mdb_file, avail_tname, act_tname)

	init do
	
		
		//Open the model database
		avail_vw = OpenTable(avail_tname, "ACCESS", {mdb_file, avail_tname})
		
		//Get list of available years
		DataYearList = V2A(GetDataVector(avail_vw+"|", "AvailYear", ))
		CloseView(avail_vw)
		DataYearList = SortArray(DataYearList)
		
		//Determine current year
		act_vw = OpenTable(act_tname, "ACCESS", {mdb_file, act_tname, "ID"})
		GetFirstRecord(act_vw+"|", )
		DataInd = ArrayPosition(DataYearList, {act_vw.Year}, )
		
		//Set to the first item in the list if not found
		if DataInd = 0 then do
			ShowMessage("Error in database setup: Invalid active year.\nPlease open the database in Access and check the active year setting.")
			Return()
		end
	
	enditem
	
	text "Select data scenario to edit" 1, 1
	
	popdown menu "Data Year"    10, 3   prompt: "Data"    List: DataYearList Variable: DataInd
	
	Button "OK" 15, 6, 10, 1.5 do
		act_vw.Year = DataYearList[DataInd]
		Return(DataYearList[DataInd])
	enditem
	
	Button "Cancel" 26.5, 6, 10, 1.5 do
		Return()
	enditem


EndDbox

Macro "CloseSelect"
    shared SelectMap
    SelectMap.MP = null
    on NotFound goto DashClosed
    UpdateDbox("Dashboard")
    DashClosed:
EndMacro