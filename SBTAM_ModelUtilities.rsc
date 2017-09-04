//******************************************************************************
//**                                                                          **
//**          LSA Associates Travel Model Utility Functions                   **
//**                                                                          **
//**          The Utilities object contains general-purpose                   **
//**   utilties useful in the implementation of travel models in TransCAD.    **
//**                                                                          **
//**       Version 2.0 - Designed for TransCAD 5.0                            **
//**                                                                          **
//**       NOTE: This version is tailored for the SLO City Model              **
//**                                                                          **
//**                                                                          **
//** ------------------------------------------------------------------------ **
//** Search for the string 'Class "' to locate the start of each object       **
//**                                                                          **
//** File Contents:                                                           **
//**  - Utilities: General purpose utilities                                  **
//**  - TripgenUT: Sociodata / Trip Generation utilities                      **
//**  - TransitUT: Route System/Transit Utilities                             **
//**                                                                          **
//******************************************************************************



Class "Utilities" //StartClass

//******************************************************************************
//** Contents:                                                                **
//**  CreateOutputDirs: Creates output directories required for a model step  **
//**  OpenView: Open a table given filename only                              **
//**  AddViewFields: Add fields to a view                                     **
//**  DropViewFields: Drops fields from a view                                **
//**  RenameViewFields: Renames fields in a view                              **
//**  CreateDSN: Create a file for MS Access database access                  **
//**  OpenDSN: Opens an Access table in a temporary FFB file                  **
//**  GetSaveAs: Gets a Save As filename from the user                        **
//**  GetOpen: Gets an Open filename from the user                            **
//**  ChooseDir: Gets a directory name from the user                          **
//**  DeleteMtxIndex: Deletes a matrix index (failsafe)                       **
//**  AddMtxCore: Adds a matrix core (failsafe)                               **
//**  IEIncices: Create Internal and External matrix indices                  **
//**  FormatDate: Converts the data and time into a nice format               **
//**  KillBars: Kills any left-over progress bars                             **
//**  Keys: Return a list of keys in an array of {key, value} pairs.          **
//**  Delete: Deletes a file, table, or geographic file                       **
//**  Values: Return a list of values in an array of {key, value} pairs.      **
//**                                                                          **
//******************************************************************************


	Macro "CreateOutputDirs" (Opts) do
	
		//Get a list of directories that need to be created
		req_dirs = null
		for i = 1 to Opts.length do
			file = Opts.(Opts[i][1]).Value  //Each file in the options array
			dir = SplitPath(file)
			dir = dir[1]+dir[2]
			
			//If this directory is not already in the list
			if ArrayPosition(req_dirs, {dir}, ) = 0 then do
				req_dirs = req_dirs + {dir}
			end
		end
		
		//Create each directory in the list
		for i = 1 to req_dirs.length do
		
			//Get all parent directory info
			tmp = SplitPath(req_dirs[i])
			drive = tmp[1]
			dir_parts = ParseString(tmp[2], "\\")
			
			cur_dir = tmp[1]  //e.g., "C:\\"
			if right(cur_dir, 1) <> "\\" then cur_dir = cur_dir + "\\"
			for j = 1 to dir_parts.length do
				cur_dir = cur_dir + dir_parts[j]  //no trailing backslash! (yet)
				if GetDirectoryInfo(cur_dir, "Directory") = null then do //GetDirectoryInfo works only w/o a trailing backslash
					CreateDirectory(cur_dir)
				end
				cur_dir = cur_dir + "\\"
			end
		end
	EndItem //EndMethod

//******************************************************************************
// OpenView: 
//  Open a table (FFB, FFA, CSV, or DBF) and use the filename as the intended
//   view name.  Return the actual view name.
	Macro "OpenView" (fname) do
    
        FT = null
        FT.BIN = "FFB"
        FT.ASC = "FFA"
        FT.DBF = "DBASE"
        FT.CSV = "CSV"
        
        t = SplitPath(fname)
        ext = Substring(Upper(t[4]), 2, )
        vw = t[3]
        
        type = FT.(ext)
        if type = null then do
            Throw("Cannot open file " + fname + "\nFile extension not recognized")
        end
        
        vw = OpenTable(vw, type, {fname})
        
        Return(vw)
        
        
    
    EndItem //EndMethod
    
    
//******************************************************************************
// AddViewFields: 
//   Add new fields to a table, don't add a duplicate if it already exists.
//   This is similar to the "TCB Add View Fields" macro, but behaves a bit
//   differently.

//   NewFlds = {{"Field1", "Integer/Real/Short/Tiny/Float", [width], [decimals]},
//              {"Field2", "Integer/Real/...", [width], [decimals]}, {...}}
//                --> [width] and [decimals] are optional integers.  Defaults 
//                    are 9 and 2 respectively.  15 and 4 are recommended 
//			          defaults for real number fields.
//   View = Name of an open, writable view
//   AfterField = Field after whcih to add new fields (optional). Fields that 
//                already exist will not be moved.
//
//                                        **********
	Macro "AddViewFields" (NewFlds, View, AfterField) do

		//Setup
		str = GetTableStructure(View)  //Load existing table structure
		dim already[NewFlds.length]    //Init variable indicating if "add" fields already exist

		//Process existing fields
		for i = 1 to str.length do
			//Check for existing field - flag if "add" field already exists
			for j = 1 to NewFlds.length do
				if str[i][1] = NewFlds[j][1] then already[j] = 1
			end
			//Set new field name to original field name on existing fields (retain existing data)
			str[i] = str[i]+{str[i][1]}
		end

		//Prepare the fields to add - only add fields that do not already exist
		modify = null
		new_str = null
		for i = 1 to NewFlds.length do
			if already[i] <> 1 then do
				w = null
				d = null
				if NewFlds[i].length >= 3 then w = NewFlds[i][3]
				if NewFlds[i].length >= 4 then	d = NewFlds[i][4]
				if w = null then w = 9
				if d = null then d = 2
				new_str = new_str + {{NewFlds[i][1], NewFlds[i][2], w, d, "False",,,,,,}}
				modify = 1
			end
		end
		
		//Add the missing fields to the str array
		if AfterField = null then  //If no AfterField was specified, add to the end
			mod_str = str + new_str
		else do
			FieldPos = null
			for i = 1 to str.length do
				if str[i][1] = AfterField then
					FieldPos = i
			end
			if FieldPos = null then //If the field was not found, add to the end
				mod_str = str + new_str
			else  //If the field is found add new fields immediately following
				mod_str = SubArray(str, null, FieldPos) + new_str + SubArray(str, FieldPos+1, )
		end

		if modify = 1 then ModifyTable(View, mod_str)
		Return(1)

	EndItem //EndMethod


//******************************************************************************
//DropViewFields: Removes fields from table, don't fail if a field doesn't exist
// - Flds: Fields to drop = {"Field1", "Field2", ...}
// - View = Name of an open, writable view
	Macro "DropViewFields" (Flds, View) do

    	//Setup and data gathering
    	str = GetTableStructure(View)
    	dim already[str.Length]

    	//For each field:
    	for i = 1 to str.length do
        	//Check for existing field
        	for j = 1 to Flds.length do
            	if str[i][1] = Flds[j] then already[i] = 1
				modify = 1
        	end
    	end

		//Keep only non-deleted fields
		str2 = null
		for i = 1 to already.Length do
			if already[i] <> 1 then do
        		str[i] = str[i]+{str[i][1]}  	//Set new field name to original field name
				str2 = str2 + {str[i]}
			end
		end
    	if modify = 1 then ModifyTable(View, str2)

    	Return(1)

	EndItem //EndMethod

//******************************************************************************
//RenameViewFields: Rename fields without changing any data or data types
// - FieldNames = {{"Old Name", "New Name"}, {"Old Name 2", "New Name 2"}, {...}}
// - View = Name of an open, writable view
	Macro "RenameViewFields" (FieldNames, View) do

    	//Setup and data gathering
    	str = GetTableStructure(View)

		for i = 1 to str.length do
			//Add an element to the end of the array, indicating that original data should be used
			str[i] = str[i] + {str[i][1]}
			//Change the name of the field if necessary
			for j = 1 to FieldNames.length do
				if FieldNames[j][1] = str[i][1] then 
					str[i][1] = FieldNames[j][2]
			end
		end

    	ModifyTable(View, str)
    	Return(1)

	EndItem //EndMethod

    
//******************************************************************************
//OpenTable: Open a table, detectign the type and creating a view named the same
//           as the table filename.  Will not open Excel or Access tables.
// - fname = full file path and name
    Macro "OpenTable" (fname, tname) do
    
        t = SplitPath(fname)
        ext = Substring(Upper(t[4]), 2, )
        vw = t[3]
        
        Types = null
        Types.DBF = "DBASE"
        Types.ASC = "FFA"
        Types.BIN = "FFB"
        Types.CSV = "CSV"
        
        type = Types.(ext)
        if type = null then Throw("Cannot open table - invalid extension\n" + fname)
        
        new_vw = OpenTable(vw, type, {fname})
        
        Return(new_vw)
    
    EndItem //EndMethod
    
//******************************************************************************
// CreateDSN: Creates a DSN file that can be used to read or write data from
// the MS Access database that is passed in the variable Filename.  The returned
// value is the name of a the created DSN file, which is created in the TransCAD
// temporary directory.

    	//Re-using DSN files results in incorrect management of database access.
		//Instead, a new temp DSN file is created each time this macro is run.
	
	Macro "CreateDSN" (Filename) do
		
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
	EndItem //EndMethod

//******************************************************************************
//OpenDSN: Open an Access table using a DSN file, the table/query's name, and an
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
//                                    *****  ***********
	Macro "OpenDSN" (dsn_name, tname, index, target_file) do

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
	EndItem //EndMethod
	
//******************************************************************************
// GetSaveAs: Gets a filename using the save-as method.  Returns null if 
//            the user cancels (instead of an error message)
// - ftype: File type description formatted as:
//          {{"Type Name", "*.ext"}, {"etc", "*.etc"}}
// - title: Dialog window title
// - init_dir: Initial file directory (Optional)
// - init_name: Suggested file name (Optional)
//
//                                   ********  *********
	Macro "GetSaveAs" (ftype, title, init_dir, init_name) do

		if right(init_dir, 1) = "\\" then  //remove trailing backslash
			init_dir = left(init_dir, StringLength(init_dir) - 1)

		Opts = {{"Initial Directory",     init_dir},
				{"Suggested Name",        init_name}}

		on escape goto GetSaveCancel
			newname = ChooseFileName(ftype, title, Opts)
		on escape default

		return(newname)

		GetSaveCancel:
		return(null)
	EndItem //EndMethod

//******************************************************************************
// GetOpen: Gets a filename using the Open method.  Returns null if 
//            the user cancels (instead of an error message)
// - ftype: File type description formatted as:
//          {{"Type Name", "*.ext"}, {"etc", "*.etc"}}
// - title: Dialog window title
// - init_dir: Initial file directory (Optional)
//
//                                 ********
	Macro "GetOpen" (ftype, title, init_dir) do
		//File Type is: {{"Type Name", "*.ext"}}

		if right(init_dir, 1) = "\\" then  //remove trailing backslash
			init_dir = left(init_dir, StringLength(init_dir) - 1)
		Opts = {{"Initial Directory",     init_dir}}

		on escape goto GetOpenCancel
			newname = ChooseFile(ftype, title, Opts)
		on escape default

		return(newname)

		GetOpenCancel:
		return(null)
	EndItem //EndMethod

//******************************************************************************
// ChooseDir: Choose a directory using the Windows dialog.  Returns null if 
//            the user cancels (instead of an error message)
// - prompt: Dialog text prompt
// - init_dir: Initial file directory (Optional)
//
//                             ********
	Macro "ChooseDir" (prompt, init_dir) do

		on escape goto canceled
		if right(init_dir, 1) = "\\" then
			init_dir = left(init_dir, StringLength(init_dir) - 1)
		ret_dir = ChooseDirectory(prompt, {{"Initial Directory", init_dir}})
        if right(ret_dir, 1) <> "\\" then ret_dir = ret_dir + "\\"
		return(ret_dir)

		canceled:
		return(null)

	EndItem //EndMethod
	
//******************************************************************************
//DeleteMtxIndex: Delete a matrix index if it exists, do nothing if it doesn't
// - mat: Matrix handle
// - idx_name: Index name
	Macro "DeleteMtxIndex" (mat, idx_name) do

		on NotFound goto DeleteMtxIndexNext
		DeleteMatrixIndex(mat, idx_name)
		DeleteMtxIndexNext:
		on NotFound default
	EndItem //EndMethod

//******************************************************************************
//AddMtxCore: Macro to add a matrix core, replacing the core if it already 
//            exists. This will not successfully replace the first core in a 
//            matrix, as it cannot be deleted. In this case, the macro will fail.

	Macro "AddMtxCore" (mat, CoreName) do

		on NotFound goto AddMtxCoreNext
		DropMatrixCore(mat, CoreName)
		AddMtxCoreNext:
		on NotFound default
		AddMatrixCore(mat, CoreName)

		Return(1)
	EndItem //EndMethod
	
//******************************************************************************
// IEIndices: Create "Internal" and "External" matrix indices
//
// m_file = matrix filename
// vw_file = view identifying external stations
// i_qry = query identifying internal zones
// e_qry = query identifying external zones
// z_id = field in view matching default matrix indices
	Macro "IEIndices" (m_file, vw_file, i_qry, e_qry, z_id) do

		//Internal
		Opts = null
		Opts.Input.[Current Matrix] = m_file
		Opts.Input.[Index Type] = "Both"
		Opts.Input.[View Set] = {vw_file, "VW", "Int", i_qry}
		Opts.Input.[Old ID Field] = {vw_file, z_id}
		Opts.Input.[New ID Field] = {vw_file, z_id}
		Opts.Output.[New Index] = "Internal"
		
		ret_value = RunMacro("TCB Run Operation",512 , "Add Matrix Index", Opts)
		if !ret_value then Return()	
		
		//External
		Opts = null
		Opts.Input.[Current Matrix] = m_file
		Opts.Input.[Index Type] = "Both"
		Opts.Input.[View Set] = {vw_file, "VW", "Ext", e_qry}
		Opts.Input.[Old ID Field] = {vw_file, z_id}
		Opts.Input.[New ID Field] = {vw_file, z_id}
		Opts.Output.[New Index] = "External"
		
		ret_value = RunMacro("TCB Run Operation",511 , "Add Matrix Index", Opts)
		if !ret_value then Return()
		
		Return(True)

	EndItem	 //EndMethod
	

//******************************************************************************
//FormatDate: Returns a nicely formatted date when passed a date/time formatted
//            as by TransCAD's GetDateAndTime() function
//
//            If no value is passed, the current date and time is returned.
	Macro "FormatDate" (daytime) do

		if daytime = null then daytime = GetDateAndTime()
		
		date_arr = ParseString(daytime, " ")
		day = date_arr[1]
		mth = date_arr[2]
		num = date_arr[3]
		time = date_arr[4]
		time_arr = ParseString(time, ":")
		year = date_arr[5]

		//Define day strings
		Days.Sun = "Sunday"
		Days.Mon = "Monday"
		Days.Tue = "Tuesday"
		Days.Wed = "Wednesday"
		Days.Thu = "Thursday"
		Days.Fri = "Friday"
		Days.Sat = "Saturday"

		//Define month strings
		Months.Jan = "January"
		Months.Feb = "February"
		Months.Mar = "March"
		Months.Apr = "April"
		Months.May = "May"
		Months.Jun = "June"
		Months.Jul = "July"
		Months.Aug = "August"
		Months.Sep = "September"
		Months.Oct = "October"
		Months.Nov = "November"
		Months.Dec = "December"

    	//Format the date string
    	if      s2i(time_arr[1]) = 0 then  RetVal = Days.(day) + ", " + Months.(mth) + " " + num + ", " + year + " (" + "12"                     + ":" + time_arr[2] + " AM)"
		else if s2i(time_arr[1]) < 12 then RetVal = Days.(day) + ", " + Months.(mth) + " " + num + ", " + year + " (" + time_arr[1]              + ":" + time_arr[2] + " AM)"
		else if s2i(time_arr[1]) = 12 then RetVal = Days.(day) + ", " + Months.(mth) + " " + num + ", " + year + " (" + time_arr[1]              + ":" + time_arr[2] + " PM)"
		else                               RetVal = Days.(day) + ", " + Months.(mth) + " " + num + ", " + year + " (" + i2s(s2i(time_arr[1])-12) + ":" + time_arr[2] + " PM)"

		//Return
		Return(RetVal)
	EndItem //EndMethod
	
//******************************************************************************
// KillBars: Closes any left-over progress bars
	Macro "KillBars" do
		on notfound do
				keepgoing = 0
				goto KillBarsNext
			end
			FoundBar = 0
			keepgoing = 1
			bars = 0
			while keepgoing = 1 and bars < 5 do
				DestroyProgressBar()
				FoundBar = 1
				bars = bars + 1
			end
		KillBarsNext:
		on notfound default
		if FoundBar = 1 then DisableProgressBar()

		//Reset status bar title to "Status"
		EnableProgressBar("Status", 2)
		DisableProgressBar()
	EndItem //EndMethod
	
//******************************************************************************
// Delete: Delete file utility
//  This macro deletes the named file and then returns the filename.
//  - If the named file has extension ".bin" then the macro deletes the 
//    .DCB dictionary if it is present.
//  - If the named file has extension ".dbd" then the macro attempts to delete
//    database files
//    
//
//  This is used to clear old files in macro initialization.

	Macro "Delete" (del_file) do
	
		if GetFileInfo(del_file) <> null then do
			tmp = SplitPath(del_file)
			
			//if a database, delete all associated files
			if tmp[4] = ".dbd" then do
				ext = {".dbd", ".ipx", ".sty", ".r1", ".r0", ".pts", ".pnk", ".lok", 
				       ".grp", ".dsk", ".des", ".dcb", ".cdd", ".bx", ".bin"}
				n_ext = {".sty", ".dcb", ".bx", ".bin"}
				
				for i = 1 to ext.length do
					file = tmp[1]+tmp[2]+tmp[3]+ext[i]
					if GetFileInfo(file) <> null then DeleteFile(file)
				end
				for i = 1 to n_ext.length do
					file = tmp[1]+tmp[2]+tmp[3]+"_"+n_ext[i]
					if GetFileInfo(file) <> null then DeleteFile(file)
				end
			end
			
			//If a FFB Table
			else if tmp[4] = ".bin" then do
				dcb_file = tmp[1]+tmp[2]+tmp[3]+".DCB"
				if GetFileInfo(dcb_file) <> null then DeleteFile(dcb_file)
			end
			
			//any other file type
			else do
				DeleteFile(del_file)
			end
		end
		
		Return(del_file)
		
	EndItem //EndMethod
    //**************************************************************************
    //** Return a list of keys in an array of {key, value} pairs.
    Macro "Keys" (var) do
    //** Opts array var = key/value pairs
    //**
    //** Returns --> Array r = list of keys
    //**************************************************************************
    
        if TypeOf(var) != 'array' then do
            Throw("Invalid argument passed to UT.Keys")
        end

        dim r[var.length]
        for i = 1 to var.length do
            r[i] = var[i][1]
        end
        
        var = null
        Return(r)
        
    //EndMethod
    EndItem
    
    //**************************************************************************
    //** Return a list of values in an array of {key, value} pairs.
    Macro "Values" (var) do
    //** Opts array var = key/value pairs
    //**
    //** Returns --> Array r = list of values
    //**************************************************************************
    
        if TypeOf(var) != 'array' then do
            Throw("Invalid argument passed to UT.Values")
        end

        dim r[var.length]
        for i = 1 to var.length do
            r[i] = var[i][2]
        end
        
        var = null
        Return(r)
        
    //EndMethod
    EndItem
    
    //**************************************************************************
    //** Returns a list of the select query names in a .qry file
    Macro "SelectList" (qry_file) do
    //** String qry_file = path and filename to select link query file
    //** 
    //** Returns -> Array qry_list = list of query names in the file
    //**
    //** NOTE: This is compatible with the query file format generated by 
    //**       TransCAD 5, but is **NOT** fully XML compliant.  This simple
    //**       appraoch simply looks for <name>****</name> and returns a list
    //**       of identified names.  The name tags and name CAN NOT be split 
    //**       across lines.
    //**************************************************************************
    
    //check for a select query file 
    if GetFileInfo(qry_file) = null then do
        Throw("No select query file found for this scenario")
    end
    
    fp = OpenFile(qry_file, "r")
    lines = ReadArray(fp)
    CloseFile(fp)
    
    StartTag = '<name>'
    EndTag = '</name>'
    
    qry_list = null
    for i = 1 to lines.length do
        l = lines[i]
        pos = Position(l, StartTag)
        if pos > 0 then do
            pos = pos + Len(StartTag)
            epos = PositionFrom(pos, l, EndTag)
            name = Substring(l, pos, epos-pos)
            qry_list = qry_list + {name}
        end
    end
    
    Return(qry_list)
    
    EndItem
    //EndMethod

EndClass  //end class "Utilities"


//** TripgenUT Contents:                                                      **
//**  SocioBV: computes TAZ-leve bivariate distrubtion of households          **    TG
//**  RegionalMedian: Get regional median income based on TAZ median incomes  **    TG
//**  Fratar3D: Fratar a 2D or 3D marray                                      **    TG
//**  CrossClass2d: Cross Class Production Utility                            **    TG
//**  WalkAccess: Add walk access links to a network                          **    TR
//**                                                                          **
//**                                                                          **
//******************************************************************************

Class "TripgenUT" //StartClass

//Socioeconomic Data Disaggregation Utility
//
// Computes univariate distribution of households based on average or median
// value along with a series of simultaneous curves.
//
// V_distvar = distribution vector (e.g., average or median value for each zone)
// V_tothh   = total households vector
// lookup_vw = View name of an open view containing the disaggregation model 
//             (i.e., table defining simultaneous curves)
//			   Table field names are specified in the Opts array.
// Opts
//   .Index = (String) = Field name with index variable (e.g., Average or Median value)
//   .Categories = (array of strings) = Array of field names with distribution values
//   .MinValue = Minimum range of distribution value (defaults to lookup table minimum if null or out of range)
//   .MaxValue = Maximum range of distribution value (defaults to lookup table maximum if null or out of range)
//   .DefaultValue = Value to replace out of range variables 
//                   (defaults to HH-weighted average if null or out of range. Uses lookup midpoint if out of lookup range)
//                   
//
// Output: 
//  HHDist = an array of vectors containing households distributed into categories
//
// Notes:
// - All vectors must be of the same length and must be ordered consistently


	Macro "SocioDisagg" (V_distvar, V_tothh, lookup_vw, Opts ) do
	
		//1) check for the required options
		if Opts = null then do
			Throw("Required argument -- Options Array is not present")
			Return()
		end
		if Opts.Index = null | Opts.Categories = null then do
			Throw("Required argument -- Options Array is not complete")
			Return()
		end
		
		//2) check to see if total households and distribution variable vectors of same length
		if V_tothh.length <> V_distvar.length then do
			Throw("Unequal TAZ and " + Opts.Index + "vectors")
			Return()
		end
		
		//3) check for the table structure
		struct = GetViewStructure(lookup_vw)
		dim vw_flds[struct.Length]
		for i = 1 to vw_flds.length do
			vw_flds[i] = struct[i][1]
		end
		if ArrayPosition(vw_flds, {Opts.Index}, ) = 0 then Throw("Error: Index field "+Opts.Index +" not found in view " + lookup_vw)
		for i = 1 to Opts.Categories.Length do
			if ArrayPosition(vw_flds, {Opts.Categories[i]}, ) = 0 then Throw("Error: Field "+Opts.Categories[i] +" not found in view " + lookup_vw)
		end
		
		//Get lookup table index values
		INDEX = GetDataVector(lookup_vw+"|", Opts.Index, )
		
		//Define range limits
		ind_min = VectorStatistic(INDEX, "Min", )
		ind_max = VectorStatistic(INDEX, "Max", )
		if ind_min = null | ind_max = null | ind_min = ind_max then Throw("Error in lookup table")

		//Min
		if Opts.MinValue = null then MinValue = ind_min
		else if Opts.MinValue < ind_min then MinValue = ind_min
		else if MinValue >= ind_max then MinValue = ind_min
		else MinValue = Opts.MinValue
        //Max
		if Opts.MaxValue = null then MaxValue = ind_max
		else if Opts.MaxValue > ind_max then MaxValue = ind_max
		else if Opts.MaxValue <= ind_min then MaxValue = ind_max
		else MaxValue = Opts.MaxValue		
		
		if MinValue >= MaxValue then Throw("Error in Options MinValue and MaxValue")
		
		//Define default index value
		ind_avg = VectorStatistic(V_tothh * V_distvar, "Sum", ) / VectorStatistic(V_tothh, "Sum", )
		if ind_avg < MinValue | ind_avg > MaxValue then ind_avg = VectorStatistic(INDEX, "Median", )  //failsafe
		if Opts.DefaultValue = null then DefaultValue = ind_avg
		else if Opts.DefaultValue < MinValue | Opts.DefaultValue > MaxValue then DefaultValue = ind_avg
		else DefaultValue = ind_avg
		
		//defining the variables
		if V_tothh.length > 0 then num_zns = V_tothh.length
		else Throw("Total Households Vector is null")
		
		dim LowDist[Opts.Categories.length]
		dim HghDist[Opts.Categories.length]
		dim FinDist[Opts.Categories.length]
		dim HHDist[Opts.Categories.length]
		
		//building arrays of vectors
		for a = 1 to FinDist.length do
			v = Vector(num_zns, "Double", )
			FinDist[a] = v
			HHDist[a] = v
		end
		
		//Start report file section
		AppendToReportFile(0, "", {{"Section", "True"}, {"Title", "Household Disaggregation Model"}})
		AppendTableToReportFile({  {{"Name", "Name"}, {"Percentage Width", 25}, {"Alignment", "Left"}}, 
		                           {{"Name", "Value"}, {"Percentage Width", 75}, {"Alignment", "Left"}, {"Decimals", 2}}}, 
								   {{"Title", "Function Settings"}})
		AppendRowToReportFile({"Number of Zones", num_zns}, )
		AppendRowToReportFile({"Index Variable", Opts.Index}, )
		AppendRowToReportFile({"Minimum Range", MinValue}, )
		AppendRowToReportFile({"Maximum Range", MaxValue}, )
		AppendRowToReportFile({"Default Value", DefaultValue}, )
		for i = 1 to Opts.Categories.length do
			AppendRowToReportFile({"Category " + format(i, "*."), Opts.Categories[i]}, )
		end
	
		
		//interpolating the % of distributions based on the variable values
		Warn = False
		for I = 1 to num_zns do
			//Get value
			hh_var = V_distvar[I]
			if hh_var < MinValue | hh_var > MaxValue then do
				if Warn = False then do
					AppendTableToReportFile({{{"Name", "Row"}}, {{"Name", "Value"}, {"Decimals", 2}}, {{"Name", "Replacement Value"}, {"Decimals", 2}}}, {{"Title", "Out of Range Errors Encountered"}})
					Warn = True
				end
				//AppendRowToReportFile({"Name", Format(I, "*.") + " is out of range (" + Format(hh_var, "*.00") + "). Using Default."}, )
				AppendRowToReportFile({I, hh_var, DefaultValue}, )
				hh_var = DefaultValue
			end
			rnd_val = Round(hh_var, 2)
			
			//get floor
			flr_var = Floor(rnd_val*10)/10
			flr_rec = LocateRecord(lookup_vw +"|", Opts.Index, {flr_var}, {{"Exact", "True"}})
			for j = 1 to LowDist.length do
				tmp = Opts.Categories[j]
				LowDist[j] = lookup_vw.(tmp)
			end
			
			//get ceil
			cel_var = Ceil(rnd_val*10)/10
			cel_rec = LocateRecord(lookup_vw +"|", Opts.Index, {cel_var}, {{"Exact", "True"}})
			for j = 1 to HghDist.length do
				tmp = Opts.Categories[j]
				HghDist[j] = lookup_vw.(tmp)
			end
			
			//interpolate the values
			for j = 1 to FinDist.length do
				if(cel_var - rnd_val =0) then do
					FinDist[j][I] = HghDist[j]
				end
				else do
					//FinDist[j][I] = (((rnd_val - flr_var)*(HghDist[j] - LowDist[j]))/(cel_var - rnd_val)) + LowDist[j]
					FinDist[j][I] = ((HghDist[j] - LowDist[j]) / (cel_var - flr_var)) * (rnd_val - flr_var) + LowDist[j]
				end
			end
		end
		
		//normalizing the distributions to 100 % 
		V_err = Vector(num_zns, "Double", {{"Constant", 0}})
		V_one = Vector(num_zns, "Double", {{"Constant", 1}})
		V_sum = Vector(num_zns, "Double", {{"Constant", 0}})
		//V_sum2 = Vector(num_zns, "Double", {{"Constant", 0}})
		
		//summing the distribution for each taz
		for i = 1 to FinDist.length do
			V_sum = nz(V_sum) + nz(FinDist[i])
		end
		
		//making the error correction / normalizing
		for i = 1 to FinDist.length do
			//V_err = V_one - V_sum
			//FinDist[i] = ((FinDist[i] / V_sum)* V_err ) + FinDist[i]
			FinDist[i] = FinDist[i] / V_sum
		end
		
		//temporary testing to see if the normailzing is working or not
		/*for m = 1 to FinDist.length do
			V_sum2 = nz(V_sum2) + nz(FinDist[m])
		end	*/

		//applying the normalized distributions to total households 
		for m = 1 to FinDist.length do
			HHDist[m] = (FinDist[m]) * (V_tothh)
		end
		
		CloseReportFileSection()
		Return(HHDist)
	EndItem //EndMethod
	
	
	//Soicioeconomic Data Bivariate Distribution Utility
	//
	// Computes TAZ-level bi-variate distributions of households
	//
	// reg_bv = (2D array) = Regional bivaraite distribution of households by two variables
	// UV1 = (array of vectors) = Univariate distribution of households by variable 1 for each TAZ
	// UV2 = (array of vectors) = Univariate distribution of households by variable 2 for each TAZ
	// TOT_HH = (Vector) = Total households by TAZ
	// TAZ = (Vector) = TAZ IDs (Optional) for use in warning messages
	// 
	//Output:
	// Z_BV = (2D array of vectors) = Bivariate distribution of households by two variables for each TAZ
	//
	//Notes:
	// - All vectors must be of the same length and must be ordered consistently
	// - the TOT_HH variables is used as a control. All other inputs are used 
	//   as distributions only.
	// - This method is dependent on the Fratar3D method, which must be present in 
	//   the same class as this method.
	//
	
	Macro "SocioBV" (reg_bv, UV1, UV2, TOT_HH, TAZ) do
	
		shared canned
		
		//Define Fratar convergence options
		FratOpts = null
		FratOpts.MaxIter = 100
		FratOpts.MaxError = 0.001
	
		//Check vector length
		zone_count = TOT_HH.Length
		if zone_count = 0 or zone_count = null then Throw("Error: zero length TAZ vector")
		for i = 1 to UV1.Length do
			if UV1[i].Length <> zone_count then Throw("Error: Inconsistent input vector length (UV1)")
		end
		for j = 1 to UV2.Length do
			if UV2[j].Length <> zone_count then Throw("Error: Inconsistent input vector length (UV2)")
		end
		if TAZ <> null then if TAZ.Length <> zone_count then Throw("Error: Inconsistent input vector length (TAZ)")
		
		//Check bivariate consistency
		if UV1.Length <> reg_bv.length then Throw("Error: Inconsistent bivariate data specification (UV1)")
		for i = 1 to reg_bv.length do
			if UV2.Length <> reg_bv[i].Length then Throw("Error: Inconsistent bivariate data specification (UV2)")
		end
		
		AppendToReportFile(0, "", {{"Section", "True"}, {"Title", "Bivariate Fratar Model"}})
		AppendTableToReportFile({  {{"Name", "Name"}, {"Percentage Width", 25}, {"Alignment", "Left"}}, 
		                           {{"Name", "Value"}, {"Percentage Width", 75}, {"Alignment", "Left"}, {"Decimals", 10}}}, 
								   {{"Title", "Function Settings"}})
		AppendRowToReportFile({"Number of Zones", TAZ.Length}, )
		AppendRowToReportFile({"Dimensions", Format(UV1.Length, "*.") + " x " + Format(UV2.Length, "*.")}, )
		AppendRowToReportFile({"Maximum Iterations", FratOpts.MaxIter}, )
		AppendRowToReportFile({"Convergence", FratOpts.MaxError}, )
		
		
		//copy input arrays to prevent overwriting input data
		reg_bv = CopyArray(reg_bv)
		UV1 = CopyArray(UV1)
		UV2 = CopyArray(UV2)
		
		//Convert regional bivariate to % distribution
		s = 0
		for i = 1 to reg_bv.Length do
			for j = 1 to reg_bv[i].Length do
				s = s + nz(reg_bv[i][j])
			end
		end
		for i = 1 to reg_bv.Length do
			for j = 1 to reg_bv[i].Length do
				reg_bv[i][j] = reg_bv[i][j] / s
			end
		end
		
		//convert univariate marginals to % distribution
		UV1_s = 0
		for i = 1 to UV1.Length do
			UV1_s = UV1_s + UV1[i]
		end
		for i = 1 to UV1.Length do
			UV1[i] = nz(UV1[i] / UV1_s)
		end
		
		UV2_s = 0
		for j = 1 to UV2.Length do
			UV2_s = UV2_s + UV2[j]
		end
		for j = 1 to UV2.Length do
			UV2[j] = nz(UV2[j] / UV2_s)
		end
		
		//Initialize output array (i,j,Z)
		dim Z_BV[UV1.Length, UV2.Length, TAZ.Length]
		
		//Fratar income and size marginals to match regional marginals

		//Seed data - convert to arrays
		for i = 1 to UV1.length do
			UV1[i] = VectorToArray(UV1[i])
		end
		for j = 1 to UV2.Length do
			UV2[j] = VectorToArray(UV2[j])
		end

		//income, size = column marginals
		dim m1_inc[UV1.Length]
		dim m1_sz[UV2.Length]
		for i = 1 to UV1.length do
			for j = 1 to UV2.length do
				m1_inc[i] = nz(m1_inc[i]) + reg_bv[i][j]
				m1_sz[j] = nz(m1_sz[j]) + reg_bv[i][j]
			end
		end
		
		//Total households = column marginal
		m2_hh = VectorToArray(TOT_HH)
		
		//Fratar income
		FR = self.Fratar3D(UV1, m1_inc, m2_hh, null, CopyArray(FratOpts))
		UV1 = FR[1]
		conv = FR[2]
		Iter = FR[3]
		if conv = False then AppendTableToReportFile({{{"Name", "WARNING: Regionwide Income Fratar did not converge."}}}, )
		AppendTableToReportFile({{{"Name", "Regional Income Iterations."}, {"Width", 300}}, 
		                         {{"Name", Format(Iter, "*.")}, {"Width", 100}}}, )
		
		//Fratar size
		FR = self.Fratar3D(UV2, m1_sz, m2_hh, null, CopyArray(FratOpts))
		UV2 = FR[1]
		conv = FR[2]
		Iter = FR[3]
		if conv = False then AppendTableToReportFile({{{"Name", "WARNING: Regionwide Income Fratar did not converge."}}}, )
		AppendTableToReportFile({{{"Name", "Regional Size Iterations."}, {"Width", 300}}, 
		                         {{"Name", Format(Iter, "*.")}, {"Width", 100}}}, )


		//Fratar each TAZ
		TAZ_Warn = False
		dim m1[UV1.length]
		dim m2[UV2.length]
		CreateProgressBar("TAZ Bivariate Seeds", "True")
		for Z = 1 to zone_count do
			canned = UpdateProgressBar("TAZ Bivariate Seeds", r2i(Z / zone_count * 100))
			if canned then return(False)
			
			//run the Fratar
			for i = 1 to UV1.length do
				m1[i] = UV1[i][Z]
			end
			for j = 1 to UV2.length do
				m2[j] = UV2[j][Z]
			end
			FR = self.Fratar3D(reg_bv, m1, m2, null, CopyArray(FratOpts))
			bv = FR[1]
			conv = FR[2]
			Iter = FR[3]
			
			
			Verbose = False   //Setting to report extra information about TAZ convergence
			if Verbose = True then do
				if Z = 1 then do
					AppendTableToReportFile({{{"Name", "TAZ"}}, {{"Name", "Iterations"}}, {{"Name", "Convergence"}}}, {{"Title", "TAZ Fratar Results (Verbose)."}})
				end //if
				
				if conv then conv_text = "Yes"
				else conv_text = "No"
				
				AppendRowToReportFile({TAZ[Z], conv_text, Iter}, )
			end //end verbose - move on to non-verbose warnings
			//Warn non-convergence
			else if conv = False then do
				if TAZ_Warn = False then do
					AppendTableToReportFile({{{"Name", "TAZ"}}}, {{"Title", "Some Traffic Analysis Zones did not converge."}})
					TAZ_Warn = True
				end
				AppendRowToReportFile({TAZ[Z]}, )
			end
					
			
			//Put the results into the Z_BV array (one zone at a time...)
	        for i = 1 to UV1.Length do
				for j = 1 to UV2.Length do
					Z_BV[i][j][Z] = bv[i][j]
				end
			end
		end
		DestroyProgressBar()
		
		
		//2D Fratar to regional total
		dim big_seed[UV1.Length*UV2.Length]
		dim big_m1[UV1.Length*UV2.Length]
		
		//Create a big seed and big marginals
		k = 0
		for i = 1 to UV1.length do
			for j = 1 to UV2.length do
				k = k + 1
				big_seed[k] = Z_BV[i][j]
				big_m1[k] = reg_bv[i][j]
			end
		end
		big_m2 = VectorToArray(TOT_HH)
		
		//run the big Fratar
		FR = self.Fratar3D(big_seed, big_m1, big_m2, null, FratOpts)
		bv = FR[1]
		conv = FR[2]
		Iter = FR[3]
		
		if conv = False then AppendTableToReportFile({{{"Name", "WARNING: Regionwide Fratar did not converge."}}}, )
		AppendTableToReportFile({{{"Name", "Regional Iterations."}, {"Width", 300}}, 
		                         {{"Name", Format(Iter, "*.")}, {"Width", 100}}}, )
		
		//Put results into the Z_BV array, convert arrays to vectors while we're at it
		k = 0
		for i = 1 to UV1.length do
			for j = 1 to UV2.length do
				k = k + 1
				Z_BV[i][j] = ArraytoVector(bv[k], {{"Type", "Double"}})
			end
		end    
		
		//complete the report
		CloseReportFileSection()
		//Return the result
		Return(Z_BV)
	
	EndItem //EndMethod
	
	
	//Regional Median Income calculation utility
	//
	// Estimates the regional median income based on TAZ median incomes
	//
	// TotHH = Vector of total households in each TAZ
	// MedInc = Vector of median income for each TAZ
	//
	//Output:
	// Regional median number (single number)
	//
	//Notes
// - All vectors must be of the same length and must be ordered consistently
	
	Macro "RegionalMedian" (TotHH, MedInc) do
	
	
		//error checks to be done
		//1) check to see if all the three vectors passed to this macro are of same size
		if TotHH.length <> MedInc.length then do
			Throw("Household and Income Vectors are not of the same length")
			Return()
		end
		
		//2) also, zones with median income but 0 households or zones with households but 0 median income
		//... not currently checked ...

		//summing all households in the model
		all_HH = VectorStatistic(TotHH, "Sum", )
		//half way between the total number of households
		hlf_HH = (all_HH*0.5) + 0.5

		//sort all the vectors based on Median Income
		//assumes that all the vectors are sent in the same order to this macro.
		srt_A = SortVectors({MedInc, TotHH}, )
		
		//calculating  running sum of the households and 
		sum_hh = 0
		fact = 0
		med_inc = 0 
		for i = 1 to srt_A[1].length do
			sum_hh = sum_hh + srt_A[2][i]
			if sum_hh => hlf_HH then do
				low_inc = srt_A[1][i-1]
				hgh_inc = srt_A[1][i]
				fact = (hlf_HH - (sum_hh - srt_A[2][i]))/ srt_A[2][i]
				med_inc = (fact * (hgh_inc - low_inc)) + low_inc
				//ShowArray({half_HH, sum_hh, hgh_inc, low_inc})
				Return(med_inc)
			end
		end
		Throw("Error computing regional median income")
	EndItem //EndMethod


//3D Fratar utility
//
// seed: n x m x p array of fratar seed data (or a 2D seed)
// m1: marginal array of length n
// m2: marginal array of length m
// m3: marginal array of length p  (null for 2D fratar)
// Opts:
//   .MaxIter: Maximum iterations (defaults to 100)
//   .MaxError: Error stopping criteria (defaults to 0.001)
//
// Note: The iteration ALWAYS stops on the last marginal.
//       2D fratar can also be performed by this function.
//
// Returns:
// {Result, Converge, Iter}
//   --> Converge is True if converged, False if forced to stop
//   --> Iter is the number of iterations required to converge
	
	Macro "Fratar3D" (seed, m1, m2, m3, Opts) do
	
		//Check options
		MaxIter = Opts.MaxIter
		if MaxIter = null then MaxIter = 100
		MaxError = Opts.MaxError
		if MaxError = null then MaxError = 0.001
		
		//Add backwards compatibility for 2D fratar
		if m3 = null then do
			F_2D = True
			for i = 1 to m1.length do
				for j = 1 to m2.length do
					typ = TypeOf(seed[i][j])
					if !(typ = "int" | typ = "double" | typ = "null") then Throw("Error: Fratar dimension inconsistency")
				end
			end
			seed = {CopyArray(seed)}
			m3 = CopyArray(m2)
			m2 = CopyArray(m1)
			m1 = {1}
		end
		else F_2D = False
		//End add backwards compatibility for 2D fratar
		
		//Check input integrity
		if seed.Length <> m1.Length then Throw("Error: misaligned FRATAR inputs")
		for i = 1 to m1.Length do
			if seed[i].Length <> m2.Length then Throw("Error: misaligned FRATAR inputs")
			for j = 1 to seed[i].Length do
				if seed[i][j].Length <> m3.Length then Throw("Error: misaligned FRATAR inputs")
			end
		end
		
		//Copy arrays to prevent modifying inputs
		seed = CopyArray(seed)
		m1 = CopyArray(m1)
		m2 = CopyArray(m2)
		m3 = CopyArray(m3)
		
		//Set up working vars
		dim f1[m1.length]  //factor1
		dim f2[m2.length]  //factor2		
		dim f3[m3.length]  //factor3
		
		//Normalize marginal totals to m3
		s1 = 0
		s2 = 0
		s3 = 0
		for i = 1 to m1.length do
			s1 = s1 + nz(m1[i])
		end
		for j = 1 to m2.length do
			s2 = s2 + nz(m2[j])
		end
		for k = 1 to m3.length do
			s3 = s3 + nz(m3[k])
		end
		for i = 1 to m1.length do
			if s1 = 0 then m1[i] = 0
			else m1[i] = m1[i] * (s3/s1)
		end
		for j = 1 to m2.length do
			if s2 = 0 then m2[j] = 0
			else m2[j] = m2[j] * (s3/s2)
		end
		
		s1 = null
		s2 = null
		s3 = null
		
		//Iterate
		Loop = True
		Iter = 0
		Converge = True
		while Loop = True do
			Iter = Iter + 1
		
			//factor m1
			for i = 1 to m1.length do
				s = 0
				for j = 1 to m2.length do
					for k = 1 to m3.length do
						s = s + nz(seed[i][j][k])
					end
				end
				if s = 0 then f1[i] = 1  //allow for zero row/col total
				else f1[i] = m1[i] / s
				for j = 1 to m2.length do
					for k = 1 to m3.length do
						seed[i][j][k] = seed[i][j][k] * f1[i]
					end
				end
			end
			
			//factor m2
			for j = 1 to m2.length do
				s = 0
				for i = 1 to m1.length do
					for k = 1 to m3.length do
						s = s + nz(seed[i][j][k])
					end
				end
				if s = 0 then f2[j] = 1  //allow for zero row/col total
				else f2[j] = m2[j] / s
				for i = 1 to m1.length do
					for k = 1 to m3.length do
						seed[i][j][k] = seed[i][j][k] * f2[j]
					end
				end
			end
			
			//factor m3
			for k = 1 to m3.length do
				s = 0
				for i = 1 to m1.length do
					for j = 1 to m2.length do
						s = s + nz(seed[i][j][k])
					end
				end
				if s = 0 then f3[k] = 1  //allow for zero row/col total
				else f3[k] = m3[k] / s
				for i = 1 to m1.length do
					for j = 1 to m2.length do
						seed[i][j][k] = seed[i][j][k] * f3[k]
					end
				end
			end
			
			//Check for convergence
			Loop = False
			for i = 1 to f1.length do
				if abs(1-f1[i]) > MaxError then Loop = True
			end
			for j = 1 to f2.length do
				if abs(1-f2[j]) > MaxError then Loop = True
			end
			for k = 1 to f3.length do
				if abs(1-f3[k]) > MaxError then Loop = True
			end
			
			if Loop = True and Iter >= MaxIter then do
				Converge = False
				Loop = False
			end
		end //end while Loop = True
		
		//Re-format to 2D if needed
		if F_2D then seed = CopyArray(seed[1])
		
		Return({seed, Converge, Iter})
	
	EndItem	 //EndMethod
	
//2D Cross Class Production Utility
//
// rates: 2d array of rates
// sedata: 2d array of socioeconomic data vectors (by TAZ), 
//
// Returns:
//    {tot_prod, i_prod, prod}
//   tot_prod: total productions by zone
//   i_prod: 1d array of production vectors by i (e.g., income) by TAZ - for use in 
//   prod: 2d array of productions by i and j (e.g., income and size) by TAZ
//


	Macro "CrossClass2D" (rates, sedata) do
		
		
		//Check for input data errors
		if rates.length <> sedata.length then Throw("Trip rates are not consistent with sociodata (1)")
		tmp = rates[1].length
		tmp_zones = sedata[1][1].length
		for i = 1 to rates.length do
			if rates.length <> tmp then Throw("Error in definition of trip rates")
			if sedata.length <> tmp then Throw("Trip rates are not consistent with sociodata (2)")
			for j = 1 to tmp do
				if sedata[i][j].length <> tmp_zones then Throw("Inconsistent number of TAZs in sociodata")
			end
		end
		
		dim prod[rates.length, rates[1].length] //by i (income) and j (size)
		dim i_prod[rates.length]                //by i (income) only - for segmentation
		tot_prod = 0                            //total productions
		for i = 1 to rates.length do
			for j = 1 to rates[i].length do
				prod[i][j] = nz(rates[i][j]) * nz(sedata[i][j])
				i_prod[j] = nz(i_prod[i]) + prod[i][j]
				tot_prod = nz(tot_prod) + prod[i][j]
			end
		end
		
		Return({tot_prod, i_prod, prod})
	
	EndItem //EndMethod
	
EndClass //End TripgenUT


//******************************************************************************
//** Contents:                                                                **
//**  WalkAccess: Add walk access links to a network                          **
//**  LinkRouteSys: Link, reload, and verify a route system                   **
//**                                                                          **
//******************************************************************************


Class "TransitUT" //StartClass

//******************************************************************************
// LinkRouteSys: Link a rts and dbd file, with options to reload/verify.
//   Actions Taken:
//     - Reload and verify the route system (can be disabled with options)
//        -> Displays an error message if a problem is found
//     - Check that all route stops are associated with a node in the line layer
//        -> Maps problem nodes if any are found
//
// Input:
//   rts_file: Route system to be linked
//   tdbd_file: Line layer to use for route system
//
// Options:
//  - .[Stop Node ID]: Field in the stop layer to contain node ID (required)
//  - .Silent
//     --> True: no messages will be shown unless an error occurs.
//     --> True: then the macro will not check to see if the route system
//               uses links that are not enabled.
//  - .NoReload
//     --> True: don't re-load the route system
//  - .NoVerify
//     --> True: Don't verify the route system

	Macro "LinkRouteSys" (rts_file, tdbd_file, InOpts) do
		shared scen_ui

		tag_buffer = 0.1  //Buffer used when tagging stops to nodes
		
		//Check for required options
		node_fld = InOpts.[Stop Node ID]
		if node_fld = null then do
			ShowMessage("Missing required option [Stop Node ID]")
			Return()
		end

	NextStep = "Set Route System Line Layer"
	SetStatus(1, NextStep, )

		//Verify that the stops layer was not selected as the line layer

		file_err = "False"
 		rts_split = SplitPath(rts_file)
		rts_split = rts_split[3] + "S.dbd"
		tdbd_split= SplitPath(tdbd_file)
		tdbd_split = tdbd_split[3] + tdbd_split[4]
		if rts_split = tdbd_split then file_err = "True"
		else do
			tmp = GetDbLayers(tdbd_file)
			if tmp.Length <> 2 then file_err = "True"
		end

		if file_err = "True" then do
			ShowMessage("Invalid File Selection: Cannot link route system to selected geographic file!")
			Return()
		end

		Layers = GetDBLayers(tdbd_file)
		{node_lyr, link_lyr} = Layers
		ModifyRouteSystem(rts_file, {{"Geography", tdbd_file, link_lyr}, {"Route ID", "ID"}})

	NextStep = "Load the Route System"
	SetStatus(1, NextStep, )

		tdbd_info = GetDBInfo(tdbd_file)
    	map_name = CreateMap("Route System", {{"Scope", tdbd_info[1]},{"Auto Project", "True"}})
    	lyrs = AddRouteSystemLayer(map_name, "Route System", rts_file,)
    	RunMacro("Set Default RS Style", lyrs, "TRUE", "TRUE")
    	route_lyr = lyrs[1]
    	stop_lyr  = lyrs[2]

	NextStep = "Reload the Route System"
	SetStatus(1, NextStep, )

		if !InOpts.NoReload then ReloadRouteSystem(rts_file)

	NextStep = "Verify Route System"
	SetStatus(1, NextStep, )

		if !InOpts.NoVerify then do

		/*	on Error do
				ShowMessage("Route System Verification: There is an error in the format of the route system.")
				goto quit
			end*/
			VerifyRouteSystem(rts_file, )

			on Error do
				ShowMessage("Route System Verification: One or more routes is disconnected.\n\n" +
		            		"See the TransCAD Log File for details. (Edit: Preferences, Logging)")
				goto quit
			end
			VerifyRouteSystem(rts_file, "Connected")
			on Error default
		end

	NextStep = "Tag Stops to Nodes"
	SetStatus(1, NextStep, )

		stop_cnt = GetRecordCount(stop_lyr, )
		SetDataVector(stop_lyr+"|", node_fld, Vector(stop_cnt, "Long", ), )
		missed = TagRouteStopsWithNode(route_lyr, , node_fld, tag_buffer)
		
		if !InOpts.Silent then do
		
			if missed > 0 then do
				ShowMessage("Route System Verification: " + string(missed) + " Stops are not adjacent to a node on the associated route.\n\n" +
		            		"See the selection set \"Invalid Stops\" for details.")
				SetLayer(stop_lyr)
				SelectByQuery("Invalid Stops", "Several", "Select * Where "+node_fld+" = null", )
				colors = RunMacro("G30 setup colors")
				SetIcon(stop_lyr+"|Invalid Stops", "Font Character", "Caliper Cartographic|14", 38)
				SetIconColor(stop_lyr+"|Invalid Stops", colors[5])
				SetDisplayStatus(stop_lyr+"|Invalid Stops", "Active")

				Return()  //Return without closing the map
			end
		end

		//Exit the macro if running in silent mode...
		if InOpts.Silent then do
			CloseMap(map_name)
			SetStatus(1, "@System0", )
			Return(1)
		end
			
	NextStep = "Verify Link Status"
	SetStatus(1, NextStep, )

		//Ask if the link status should be checked
		Opts = null
		Opts.Caption = "Check Links?"
		Opts.Buttons = "YesNo"
		Opts.Icon = "Question"
		ans = MessageBox("The Route System is Valid. Check for routes using disabled links?", Opts)
		if ans = "No" then do
			CloseMap(map_name)
			SetStatus(1, "@System0", )
			Return(1)
		end
		
		//If checking, ask for a network year
		SetAlternateInterface(scen_ui)
	    	Opts = null
			Opts.HideNetwork = False
			Opts.HideData = True
			year_ans = RunDbox("Set Years", null, null, tdbd_file, null, null, Opts)
		SetAlternateInterface()
		if year_ans = null then goto quit
		NetYear = year_ans[1]
		
		//Identify routes in the system
		route_names = GetRouteNames(route_lyr)
		route_links = null
		for i = 1 to route_names.length do
			tmp = GetRouteLinks(route_layer, route_names[i])
			route_links = route_links + CopyArray(tmp)
			tmp = null
		end
		
		//Separate IDs from returned info and 
		//remove duplicates
		dim link_ids[route_links.length]
		for i = 1 to link_ids.length do
			link_ids[i] = route_links[i][1]
		end
		link_ids = VectorToArray(SortVector(ArrayToVector(link_ids), {{"Unique", "True"}}))
		
		//Select route links
		SetLayer(link_lyr)
		link_cnt = SelectByIDs("RouteLinks", "Several", link_ids, )
		
		//Select route links where FT is null
		qry = "Select * where FT_"+NetYear+" = null"
		Opts = null
		Opts.[Source And] = "RouteLinks"
		cnt = SelectByQuery("Disabled Links", "Several", qry, Opts)
		
		if cnt > 0 then do
			g30_colors = RunMacro("G30 setup colors")
			SetDisplayStatus(link_lyr+"|Disabled Links", "Active")
			SetLineColor(link_lyr+"|Disabled Links", g30_colors[23])
			SetLineWidth(link_lyr+"|Disabled Links", 8)
			ShowMessage("Error.  The route system uses one or more disabled links in the selected year.  See selection set \"Disabled Links\"")
			SetStatus(1, "@System0", )
			Return()
		end
		else do
			ShowMessage("No Problem Links Found.")
			//ShowArray({link_ids, link_cnt, cnt})
			Return()
		end
			
		quit:
		CloseMap(map_name)

	SetStatus(1, "@System0", )

		Return()
	EndItem //EndMethod


	
//  WalkAccess
//  Creates walk access links
//   --> tdbd_file: Transit dbd file to receive walk links
//   --> rts_file: File to use for walk link creation (identify stop nodes)
//   --> buffer: Maximum walk link distance
//   --> max_count: maximum number of walk links per TAZ
//   --> Opts.NearNode = "FieldName" //Name of field use to tag stops to nodes
//   --> Opts.AddFields = Options to add data to the walk links, If one AddFields option 
//                        is specified, ALL must be specified and of the same length.
//     --> Opts.AddFields.Fields    = {Field1, Field2, ...} //Fields to fill and merge with existing dbd file
//     --> Opts.AddFields.Type      = {"Real", "Integer", ...}  //Type of new fields
//     --> Opts.AddFields.Method    = {"Formula/Value", "Formula/Value", ...}   //Method to fill fields referenced above
//     --> Opts.AddFields.Parameter = {x, x, ...}  //formula or value to fill field

	Macro "WalkAccess" (tdbd_file, rts_file, buffer, max_count, InOpts) do
	
	shared canned, ret_value

	UT = CreateObject("Utilities")  //Required for utilities called by this macro

	/*	shared canned, debug, ret_value
		//Set up error handlers
		if debug <> 1 then do
			on Escape do
				on Escape default
				canned = "True"
				Return()
			end
			on Error, NotFound, NonUnique, Missing, DivideByZero, EndOfFile, LanguageError do
				on Error, NotFound, NonUnique, Missing, DivideByZero, EndOfFile, LanguageError default
				ShowMessage(GetLastError({{"Reference Info", "True"}}))
				Return()
			end
		end //end definition of error handlers
		*/
		
		//Define temporary walk access dbd file
		wdbd_file = GetTempFilename(".dbd")
		
		//Identify stop near node field
		stop_nearnode = InOpts.NearNode
		
		//Open the dbd file
		RunMacro("TCB Add DB Layers", tdbd_file,,)
		Lyrs = RunMacro("TCB get DB line and node layers", tdbd_file)
		tnode_lyr = Lyrs[1]
		tlink_lyr = Lyrs[2]
		
		//Identify Transit Stop Nodes
		tmp = SplitPath(rts_file)
		stop_tfile = tmp[1]+tmp[2]+tmp[3]+"S.bin"
		stop_vw = OpenTable("TransitStops", "FFB", {stop_tfile, })
		StopNodes = GetDataVector(stop_vw+"|", stop_nearnode, )
		CloseView(stop_vw)
		
		//Get unique stop-node IDs
		StopNodes = SortVector(StopNodes, {{"Unique", "True"}})
		StopNodes = v2a(StopNodes)
		
		//Select stop-nodes
		SetView(tnode_lyr)
		SelectByIDs("AllStops", "several", StopNodes)
		AllStops = v2a(GetDataVector(tnode_lyr+"|AllStops", "ID", ))
		
		//Select TAZs
		SetView(tnode_lyr)
		SetLayer(tnode_lyr)
		SelectByQuery("Centroids", "Several", "Select * Where Zone > 0", )
		C_IDs = GetDataVector(tnode_lyr+"|Centroids", "ID", )
		
		//Empty array to hold coordinate pairs
		coord_pairs = null
		//id_pairs = null
		
		//Process each centroid
		CreateProgressBar("Walk Access (compute)", "True")
		
		//Save links in a CSV file
		csv_file = GetTempFilename(".csv")
		fp = OpenFile(csv_file, "w")
		for I = 1 to C_IDs.length do
		//for I = 1 to 50 do  //Switch to this for quick troubleshooting/testing
			canned = UpdateProgressBar("Locate Walk Access " + string(I) + " / " + string(C_IDs.length), r2i(I / C_IDs.Length * 100) )
			if canned then Return()
			
			//Identify nearby nodes
			SetLayer(tnode_lyr)
			C_coord = GetPoint(C_IDs[I])
			STOPS_rec = LocateNearestRecords(C_coord, buffer, )
			
			//Process each node
			stop_count = 0
			for j = 1 to STOPS_rec.length do
				SetRecord(tnode_lyr, STOPS_rec[j])
				STOP_id = rh2id(STOPS_rec[j])
				
				//Only add connectors to nodes with stops
				//Only add up to a certain number per centroid
					if stop_count <= max_count and ArrayPosition(AllStops, {STOP_id}, ) > 0 then do
						
					//Get stop coordinate
					SetLayer(tnode_lyr)
					S_coord = GetPoint(STOP_id)
					
					//Save coordinate pair
					WriteLine(fp, "2, " + string(C_coord.lon) + ", " + 
										string(C_coord.lat) + ", " + 
										string(S_coord.lon) + ", " + 
										string(S_coord.lat) )
					stop_count = stop_count + 1
					
				end  //end if node is a stop

			end //end loop over nearby stops
		end  //end loop over zones
		CloseFile(fp)
		DestroyProgressBar()
		
		//Close the original dbd file
		DropLayerFromWorkspace(tlink_lyr)
		DropLayerFromWorkspace(tnode_lyr)
		
		//Import the CSV file to a geographic file
		Opts = null
		Opts.Dir = null
		Opts.Geography = 1
		Opts.ID = null
		Opts.Label = "Imported"
		Opts.[Layer Name] = "Imported"
		ImportCSV(csv_file, wdbd_file, "Line", Opts)
		
		//Open the new dbd file
		RunMacro("TCB Add DB Layers", wdbd_file,,)
		Lyrs = RunMacro("TCB get DB line and node layers", wdbd_file)
		wnode_lyr = Lyrs[1]
		wlink_lyr = Lyrs[2]

		//check to see if fields should be added, filled, and merged 
		//smc 10-9-2014: Updated AddFields code block with new version that works in TC6
		if InOpts.AddFields <> null then do

            //Check for valid parameter opt (must be array of same length as # 
            //of fields)
            if InOpts.AddFields.Parameter <> null and
                      (TypeOf(InOpts.AddFields.Parameter) <> "array" or 
                      InOpts.AddFields.Parameter.length <> 
                      InOpts.AddFields.Fields.length) then do
                Throw("Walk Link Creation: " + 
                      "Option Parameter must be an array of values")
            end
        
			//Add and fill fields
            cnt = GetRecordCount(wlink_lyr, )
			dim SetVs[InOpts.AddFields.Fields.Length]
			dim Fields[InOpts.AddFields.Fields.Length]
			for i = 1 to Fields.length do
				Fields[i] = {InOpts.AddFields.Fields[i], 
                             InOpts.AddFields.Type[i]}
                if InOpts.AddFields.Parameter <> null then do
                    argtyp = InOpts.AddFields.Type[i]
                    typeconv = {{"Integer", "Long"},
                                {"Real", "Double"},
                                {"String", "String"},
                                {"Short", "Short"},
                                {"Tiny", "Short"},
                                {"Float", "Float"}}
                    typ = typeconv.(argtyp)
                    if typ = null then do
                        Throw('Walk Link Creation: Unknown "Type" option')
                    end
                    
                    SetVs[i] = {InOpts.AddFields.Fields[i], 
                                Vector(cnt, typ, 
                                {{"Constant", InOpts.AddFields.Parameter[i]}})}
                end
			end
				
			UT.AddViewFields(Fields, wlink_lyr)
            if InOpts.AddFields.Parameter <> null then 
                SetDataVectors(wlink_lyr+"|", SetVs, )
            
		end

			
		//Re-Open the geographic network
		RunMacro("TCB Add DB Layers", tdbd_file,,)
		Lyrs = RunMacro("TCB get DB line and node layers", tdbd_file)
		tnode_lyr = Lyrs[1]
		tlink_lyr = Lyrs[2]
		
		//Merge the walk links with the main database
		Opts = null
		Opts.Snap = True
		if InOpts.AddFields <> null then do
			for i = 1 to InOpts.AddFields.Fields.length do
				Opts.Fields = Opts.Fields + {{InOpts.AddFields.Fields[i], InOpts.AddFields.Fields[i]}}
			end
		end
		MergeGeography(tlink_lyr, wlink_lyr, Opts)

		
		//Close the layers
		DropLayerFromWorkspace(wlink_lyr)
		DropLayerFromWorkspace(wnode_lyr)
		DropLayerFromWorkspace(tlink_lyr)
		DropLayerFromWorkspace(tnode_lyr)
		
		//Delete the temporary walk access dbd file
		DeleteDatabase(wdbd_file)
		
		//Repair the database now that walk links have been added
		OptimizeDatabase(tdbd_file, )
		
		Return(1)
		
	EndItem //EndMethod
	
EndClass






		
