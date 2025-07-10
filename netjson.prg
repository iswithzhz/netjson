*==============================================================
* Class:        NetJson 
* Parent:       Custom
* Dependencies: NetJson.dll (C# COM Component) | ClrHost.dll(7.29)
* Environment:  Net4.61 | VFP9 7423
* Author:       ZHZ
* Description:  Uses C# component to parse JSON and create VFP objects
*               cJson ¡ú oJson | oJson ¡ú cJson | oJson ¡ú Cursor | Cursor ¡ú cJson

* cJson ¡ú oJson Type Mapping:
	* Node        ¡ú VFP Empty class          O
	* Array       ¡ú VFP Collection class     O
	* Date        ¡ú VFP date literal {^...}  D
 	* DateTime    ¡ú VFP date literal {^...}  T
	* Null        ¡ú Null                     L
	* Boolean     ¡ú .T./.F.                  L 
    * Numeric     ¡ú 123.45                   N
    * HighPrecision ¡ú 0.132456789            N
  
* oJson ¡ú Cursor Type Mapping:
	* String      ¡ú Determine C(254) or M
	* Date/Time   ¡ú Maintain original type (D/T) 
	* Object      ¡ú Convert to M type (store JSON)
	* Numeric     ¡ú Select I/N/B type
	* Numeric conversion logic:
	* 1. Numeric type (N):
	    * - If integer (laFieldInfo[m.lnField, 5] = .T.):
	        * - If field length <=9, use I (Integer) type
	        * - If field length <=18, use N(total_digits,0)
	    * - If float (has decimals):
	        * - Calculate decimal places (max 9 digits)
	        * - Total digits = integer digits + decimal digits + 1 (decimal point)
	        * - Then:
	            * - If total digits <=20, use N(total_digits, decimal_digits)
	            * - Otherwise use B (Double)
	 * 2. Other types:
	    * - D: Date       -> D NULL
	    * - T: DateTime   -> T NULL
	    * - L: Logical    -> L NULL
	    * - C: Character
	        * - If length>254, use M (Memo)
	        * - Else, calculate buffer size: max_length*1.2 (round up, min 10, max 254)
	    * - M: Memo -> M NULL
	    * - Other unknown types: V(254) NULL (variable char, max 254)
* Cursor ¡ú oJson ¡ú cJson
    * Type mapping same as above

* Created:      2025-06-17

* Version:		   
	* 1.03   Unresolved Issue 1: VFP Empty objects use hash tables internally,;
			   causing property order loss (though key-value pairs are preserved);
 			   Unresolved Issue 2: Property names converted to lowercase
                
    * 1.04   - Added toCursor() method;
             - Added Query() method with JSONPath support;
             - Added Cursor2Json() to convert tables to cJson 

* FAQ: Why use external NetJson.dll?
* 1. Precise JSON validation and error location;
	NetJson.dll strictly validates JSON against standards;
	Detects syntax errors (mismatched brackets, unquoted fields, extra commas);
	Pinpoints error locations for efficient debugging;
	Automated validation saves time with complex JSON structures.
	
* 2. High-performance parsing with lower latency;
	Optimized algorithms for large JSON data and high-frequency scenarios;
	Reduces parsing delays in real-time systems (API responses, data imports);
	Critical for WsServer.ProcessRequest() JSON handling.
	
* 3. Native VFP object support for full OOP access;
	Converts JSON directly to VFP object model;
	Enables intuitive property access (e.g., obj.name instead of manual parsing);
	Supports nested structures via Collection classes;
	Simplifies code logic with object-oriented syntax.
	
* 4. Strict type mapping for data safety;
	Ensures accurate VFP type conversion (JSON true ¡ú .T., numbers ¡ú numeric);
	Prevents runtime errors from implicit type conversions;
	Improves code maintainability with explicit typing.
*==============================================================
Define Class NetJson As Custom
	
	#DEFINE CRLF CHR(13)+CHR(10)
    
    *- =========Public Properties ============
    oNetJson     = .Null.    				&& C# parser object
    lInitialized = .F.       				&& Initialization flag
    cLastError   = ""        				&& Last error message
    
    * - C# Component ProgId
    Protected cClassName
    cClassName = "NetJson.VfpJson"
    
    Protected cDllPath
    cDllPath   = "NetJson.dll"  
    
    *- =========JsonSerializer ============
	LFormatted 		= .T.	&& Formatting
	NIndentSize 	=  2	&& Indent size

	* Reference array initial size
	Protected nRefCount, nRefCapacity
	nRefCount 		= 0
	nRefCapacity 	= 64
	Dimension HReferences[This.nRefCapacity]   && Reference tracking array
	    
    *==============================================================
    * - Method:    Init
    * - Purpose:   Initialize component
    * - Returns:   .T. - Success, .F. - Failure
    *==============================================================
    Function Init()

        Local llSuccess ;
            , loException

        m.llSuccess        = .F.
        m.loException      = .Null.
        
        This.lInitialized  = .F.
        This.cLastError    = ""
        
        * - Load C# component
        Try
            This.oNetJson = This.CreateObjects( This.cClassName , This.cDllPath )
            If Vartype( This.oNetJson ) == "O" then 
               This.oNetJson._vfp    = _vfp
               This.lInitialized    = .T.
               m.llSuccess          = .T.
            Else
                This.cLastError = "Failed to create C# component"
            Endif
        Catch To m.loException
            This.cLastError = "C# component load failed [" + This.cClassName + "]: " + ;
                              IIF( Vartype( m.loException ) == "O", loException.Message, "Unknown error" )
        Endtry
        
        If !Empty( This.cLastError ) Then
            Messagebox( This.cLastError, 16 , "NetJson" )
        Endif
        
    Endfunc

    *==============================================================
    * - Method:    Parse
    * - Purpose:   Parse JSON string (cJson -> oJson)
    * - Params:    tcJson - JSON string
    * - Returns:   Object on success, Null on failure
    *==============================================================
    Function Parse( tcJson As String ) As Object 
        * - Check initialization
        If !This.lInitialized Then 
            This.cLastError = "Component not initialized"
            Return Null
        Endif
        
        * - Validate JSON input
        If Empty( m.tcJson ) Then 
            This.cLastError = "JSON string cannot be empty"
            Return Null
        Endif
        
        * - Generate root variable name
        m.tcRootVar = "oJson" + Sys(2015)
        
        * - Validate variable name
        If !This.IsValidVarName( m.tcRootVar ) Then 
            This.cLastError = "Invalid variable name: " + m.tcRootVar
            Return Null
        Endif
        
        * - Release existing variable
        If Type( m.tcRootVar ) != "U"
           Release ( m.tcRootVar )
        Endif
        
        * - Call C# component
        Local lcErrorMsg
        m.lcErrorMsg  = ""
       
        * - Execute parsing
        m.llResult = This.oNetJson.Parse( m.tcJson , m.tcRootVar , @lcErrorMsg )
      
        * - Handle result
        If !m.llResult Then 
           This.cLastError = m.lcErrorMsg
           Return Null
        Else 
           Return &tcRootVar.
        Endif
    Endfunc

	*==============================================================
	* Serialize (Main)
	* - Purpose:   Serialize (oJson -> cJson)
	* - Params:    toObject - NetJson generated object
	* - Returns:   JSON string or empty string on error
	*==============================================================
	Function Serialize( toObject As Object) As String
		This.CLastError = ""
        This.nRefCount = 0  && Reset reference count

        * Use StringBuilder instead of string concatenation
        Local loSB
        m.loSB = CREATEOBJECT("StringBuilder")
        m.loSB.PrepareAdd(1024)  && Pre-allocate memory
		
		Local loException
        Try
            This.SerializeValue( m.toObject, m.loSB , 0 )
        Catch To m.loException
            This.CLastError = "Serialization error: " + m.loException.Message
            If !Empty( m.loException.UserValue ) Then 
                This.CLastError = This.CLastError + " | " + m.loException.UserValue
            Endif
        Endtry
        
        If !Empty( This.CLastError )
            Return ""
        Else 
            Return m.loSB.ToString()
        Endif  
	Endfunc

    *==============================================================
    * - Method:    Query
    * - Purpose:   Execute JSONPath query
    * - Params:    tcJson - JSON string
    *              tcJsonPath - JSONPath expression
    * - Returns:   JSON result string or empty on error
    *==============================================================
    Function Query(tcJson As String, tcJsonPath As String) As String
        * - Check initialization
        If !This.lInitialized
            This.cLastError = "Component not initialized"
            Return ""
        Endif
        
        * - Validate parameters
        If Empty(m.tcJson)
            This.cLastError = "JSON string cannot be empty"
            Return ""
        Endif

        If Empty(m.tcJsonPath)
            This.cLastError = "JSONPath expression cannot be empty"
            Return ""
        Endif

        Local lcResult
        lcResult = ""  && Initialize result

        Try
        * - Execute query
            m.lcResult = This.oNetJson.Query( m.tcJson, m.tcJsonPath )

        * - Handle result
            If Empty(m.lcResult)
                This.cLastError = "No matching data found"
            Else
                This.cLastError = ""  && Clear error
            Endif

        Catch To loException
        * - Handle exception
            This.cLastError = "Query error: " + Iif(Vartype(loException) == "O", ;
                			  loException.Message, "Unknown error")
            lcResult = ""  && Return empty
        Endtry

        Return m.lcResult
    Endfunc

    *==============================================================
	* - Method:    ToCursor
	* - Purpose:   Convert JSON array to VFP cursor 
	* - Params:    oCollection  - VFP collection object
	*              tcCursorName - Cursor name
	*              tnScanRows   - Rows to scan for structure (default=1, -1=all)
	* - Returns:   .T. - Success, .F. - Failure
	*==============================================================
	Function ToCursor( oCollection As Object, tcCursorName As String, tnScanRows As Integer ) As Boolean
	
		* - Handle parameter defaults
        If Pcount() < 3 Or Vartype(m.tnScanRows) != "N"
            m.tnScanRows = 1 	&& Default scan 1 row
        Else
            If m.tnScanRows = -1
               m.tnScanRows = m.oCollection.Count  && Scan all rows
            Endif 
        Endif
        
        * - Check initialization
        If !This.lInitialized Then 
            This.cLastError = "Component not initialized"
            Return .F.
        Endif
        
        * - Validate parameters
        If Vartype( m.oCollection ) != "O"
            This.cLastError = "Invalid collection object"
            Return .F.
        Endif
        
        If Empty( m.tcCursorName ) Or !This.IsValidVarName( m.tcCursorName )
            This.cLastError = "Invalid cursor name: " + m.tcCursorName
            Return .F.
        Endif
        
        * - Close existing cursor
        If Used( m.tcCursorName )
            Use In ( m.tcCursorName )
        Endif
        
		Local lnItemCount	;
			, lnIndex		;
			, loItem		;
			, lcFieldList	;
			, lcField		;
			, lvValue
				
		Local lcType		;
			, lnField		;
			, laFields[1]	;
			, lnFields		;
			, llSuccess
				
		Local Array laFieldInfo[1]  && Field structure info
		m.llSuccess = .F.  			&& Default failure
		
		Try
			* - Get item count
			m.lnItemCount = m.oCollection.Count
	        
			* - Process non-empty collection
			If m.lnItemCount > 0
				* - Calculate actual scan rows
				m.tnScanRows = Min(m.tnScanRows, m.lnItemCount)
	            
				* - Determine field structure
				m.loItem = m.oCollection.Item(1)
	            
				Do Case
					Case Vartype( m.loItem ) = "O"
						* - Get object properties
						m.lnFields = Amembers( laFields, m.loItem, 0 )
	                    
						* - Initialize field info array [field name, type, length, decimals, isInteger]
						Dimension laFieldInfo[m.lnFields, 5]
						For m.lnField = 1 To m.lnFields
							laFieldInfo[m.lnField, 1] = laFields[m.lnField]  && Field name
							laFieldInfo[m.lnField, 2] = "U"                  && Initial unknown type
							laFieldInfo[m.lnField, 3] = 0                    && Initial length
							laFieldInfo[m.lnField, 4] = 0                    && Initial decimals
							laFieldInfo[m.lnField, 5] = .T.                  && Initial integer flag
						Endfor
	                    
						* - Scan sample data to determine field properties
						Local lnScan As Integer
						For m.lnScan = 1 To m.tnScanRows
							m.loItem = m.oCollection.Item(m.lnScan)
	                        
							For m.lnField = 1 To m.lnFields
								m.lcField = laFieldInfo[m.lnField, 1]
	                            
								* Check if property exists
								If Pemstatus(m.loItem, m.lcField, 5)
									m.lvValue = Evaluate("m.loItem." + m.lcField)
									m.lcType = Vartype(m.lvValue)
	                                
									* Update field info
									Do Case
										Case m.lcType = "N"
											* Numeric type handling
											laFieldInfo[m.lnField, 2] = "N"  && Mark as numeric
	                                        
											* Check if integer
											If Int(m.lvValue) != m.lvValue
												laFieldInfo[m.lnField, 5] = .F.  && Mark as float
	                                            
												* Update decimal places
												Local lnDecimals As Integer
												m.lnDecimals = This.GetDecimalCount(m.lvValue)
												If m.lnDecimals > laFieldInfo[m.lnField, 4]
													laFieldInfo[m.lnField, 4] = m.lnDecimals
												Endif
											Endif
	                                        
											* Update integer part length
											Local lnIntLength As Integer
											m.lnIntLength = Len(Transform(Int(Abs(m.lvValue))))
											If m.lnIntLength > laFieldInfo[m.lnField, 3]
												laFieldInfo[m.lnField, 3] = m.lnIntLength
											Endif
	                                        
										Case m.lcType = "C"
											* Character type handling
											laFieldInfo[m.lnField, 2] = "C"  && Mark as character
	                                        
											* Update max length
											Local lnLength As Integer
											m.lnLength = Len(Transform(m.lvValue))
											If m.lnLength > laFieldInfo[m.lnField, 3]
												laFieldInfo[m.lnField, 3] = m.lnLength
											Endif
	                                        
										Case m.lcType = "D"
											laFieldInfo[m.lnField, 2] = "D"
										Case m.lcType = "T"
											laFieldInfo[m.lnField, 2] = "T"
										Case m.lcType = "L"
											laFieldInfo[m.lnField, 2] = "L"
										Case m.lcType = "M"
											laFieldInfo[m.lnField, 2] = "M"
										Case m.lcType = "O"
											laFieldInfo[m.lnField, 2] = "M"  && Serialize objects to memo
									Endcase
								Endif
							Endfor
						Endfor
	                    
						* - Build field list
						m.lcFieldList = ""
						For m.lnField = 1 To m.lnFields
							m.lcField = laFieldInfo[m.lnField, 1]
							m.lcType = laFieldInfo[m.lnField, 2]
	                        
							* - Create field based on info
							Do Case
								Case m.lcType = "N"  && Numeric type
									If laFieldInfo[m.lnField, 5]  && Integer
										* Integer handling
										Do Case
											Case laFieldInfo[m.lnField, 3] <= 9  && ¡Ü9 digit integer
												m.lcFieldList = m.lcFieldList + ", " + m.lcField + " I NULL"
									            
											Case laFieldInfo[m.lnField, 3] <= 18  && 10-18 digit integer
												* Total digits = integer digits + 1 (sign)
												m.lnTotal = laFieldInfo[m.lnField, 3] + 1
												m.lcFieldList = m.lcFieldList + ", " + m.lcField + " N(" + ;
									                												Transform(m.lnTotal) + ",0) NULL"
											Otherwise  && >18 digit integer
												m.lcFieldList = m.lcFieldList + ", " + m.lcField + " B(0) NULL"
										Endcase
									Else
										m.lnDecimals = MIN(laFieldInfo[m.lnField, 4], 9)
										* Total digits = integer + decimals + 1 (point) + 1 (sign)
										m.lnTotalDigits = laFieldInfo[m.lnField, 3] + m.lnDecimals + 2
								        
										* Numeric type handling (N or B)
	                                    If m.lnTotalDigits <= 20
	                                        m.lcFieldList = m.lcFieldList + ", " + m.lcField + " N(" + ;
	                                            Transform(m.lnTotalDigits) + "," + Transform(m.lnDecimals) + ") NULL"
	                                    Else
	                                        m.lcFieldList = m.lcFieldList + ", " + m.lcField + " B(" + Transform(m.lnDecimals) + ") NULL"
	                                    Endif

									Endif
	                                
								Case m.lcType = "D"
									m.lcFieldList = m.lcFieldList + ", " + m.lcField + " D NULL"
	                                
								Case m.lcType = "T"
									m.lcFieldList = m.lcFieldList + ", " + m.lcField + " T NULL"
	                                
								Case m.lcType = "L"
									m.lcFieldList = m.lcFieldList + ", " + m.lcField + " L NULL"
                                
								Case m.lcType = "C"
									* Character type handling
									Local lnMaxLength As Integer
									m.lnMaxLength = laFieldInfo[m.lnField, 3]
	                                
									If m.lnMaxLength > 254
										m.lcFieldList = m.lcFieldList + ", " + m.lcField + " M NULL"
									Else
										* Add 20% buffer, min 10 chars
										Local lnSize As Integer
										m.lnSize = Max(10, Min(254, Ceiling(m.lnMaxLength * 1.2)))
										m.lcFieldList = m.lcFieldList + ", " + m.lcField + " C(" + Transform(m.lnSize) + ") NULL"
									Endif
	                                
								Case m.lcType = "M"
									m.lcFieldList = m.lcFieldList + ", " + m.lcField + " M NULL"
	                                
								Otherwise
									* Unknown type handling
									m.lcFieldList = m.lcFieldList + ", " + m.lcField + " V(254) NULL"
							Endcase
						Endfor
	                    
						* - Clean field list
						If !Empty(m.lcFieldList)
							m.lcFieldList = Substr(m.lcFieldList, 3)
						Endif
	                    
					Otherwise
						* - Simple type handling
						* Scan samples to determine type
						Local lcSampleType As String, lnMaxLength As Integer
						m.lcSampleType = "C"
						m.lnMaxLength = 0
	                    
						* Scan sample data
						Local lnScan As Integer
						For m.lnScan = 1 To Min(m.tnScanRows, m.lnItemCount)
							m.loItem = m.oCollection.Item(m.lnScan)
							m.lcType = Vartype(m.loItem)
	                        
							* Determine strictest type
							Do Case
								Case m.lcType = "N" And Inlist( m.lcSampleType, "C", "M" )
									m.lcSampleType = "N"
								Case m.lcType = "D" And m.lcSampleType != "T"
									m.lcSampleType = "D"
								Case m.lcType = "T"
									m.lcSampleType = "T"
								Case m.lcType = "L" And m.lcSampleType = "C"
									m.lcSampleType = "L"
								Case m.lcType = "C"
									* Update max length
									m.lnLength = Len( Transform( m.loItem ) )
									If m.lnLength > m.lnMaxLength
										m.lnMaxLength = m.lnLength
									Endif
							Endcase
						Endfor
	                    
						* Create field based on sample type
						Do Case
							Case m.lcSampleType = "N"
								m.lcFieldList = "value N NULL"
							Case m.lcSampleType = "D"
								m.lcFieldList = "value D NULL"
							Case m.lcSampleType = "T"
								m.lcFieldList = "value T NULL"
							Case m.lcSampleType = "L"
								m.lcFieldList = "value L NULL"
							Case m.lcSampleType = "C" Or m.lcSampleType = "M"
								If m.lnMaxLength > 254
									m.lcFieldList = "value M NULL"
								Else
									Local lnSize As Integer
									m.lnSize = Max(10, Min(254, Ceiling(m.lnMaxLength * 1.2)))
									m.lcFieldList = "value C(" + Transform(m.lnSize) + ") NULL"
								Endif
							Otherwise
								m.lcFieldList = "value C(254) NULL"
						Endcase
				Endcase
				* - Create cursor structure
				Create Cursor ( m.tcCursorName ) ( &lcFieldList )
	            
				* - Prepare data array
				Local Array laData[1]  && Temporary data storage
	            
				* - Populate cursor data
				For m.lnIndex = 1 To m.lnItemCount
					m.loItem = m.oCollection.Item( m.lnIndex )
	                
					* - Handle different data types
					Do Case
						Case Vartype( m.loItem ) = "O"
							* - Object type handling
							m.lnFields = Amembers( laFields, m.loItem, 0 )
							Dimension laData(m.lnFields)
	                        
							* - Traverse object properties
							For m.lnField = 1 To m.lnFields
								m.lcField = laFields[m.lnField]
	                            
								* - Check if property exists
								If Pemstatus(m.loItem, m.lcField, 5)  && 5 = property exists
									m.lvValue = Evaluate( "m.loItem." + m.lcField )
	                                
									* - Handle nested objects/collections
									Do Case
										Case Vartype( m.lvValue ) = "O" And Pemstatus( m.lvValue, "BaseClass", 5 ) And ;
													Lower( m.lvValue.BaseClass ) == "collection"
											* - Serialize nested collection to JSON
											laData[m.lnField] = This.Serialize( m.lvValue )
										Case Vartype( m.lvValue ) = "O"
											* - Serialize nested object to JSON
											laData[m.lnField] = This.Serialize( m.lvValue )
										Case Vartype( m.lvValue ) = "C" And Len( m.lvValue ) > 254
											* - Long string handling
											laData[m.lnField] = m.lvValue  && Store directly to memo
										Otherwise
											* - Basic types assign directly
											laData[m.lnField] = m.lvValue
									Endcase
								Else
									* - Non-existent properties set to NULL
									laData[m.lnField] = .NULL.
								Endif
							Endfor
	                        
							* - Insert data into cursor
							Insert Into ( m.tcCursorName ) From Array laData
	                        
						Otherwise
							* - Simple type handling
							Dimension laData(1)
							laData[1] = m.loItem
	                        
							* - Insert data into cursor
							Insert Into ( m.tcCursorName ) From Array laData
					Endcase
				Endfor
	            
				* - Go to top of cursor
				Go Top In ( m.tcCursorName )
	            
				This.cLastError = ""
				m.llSuccess 	= .T.  && Mark success
	            
			Else
				* - Empty collection handling
				Create Cursor ( m.tcCursorName ) ( value C(254) NULL )
				This.cLastError = "Empty collection"
				m.llSuccess 	= .T.  && Considered success
			Endif
	        
		Catch To m.loException
			This.cLastError = "ToCursor error: " + IIF( Vartype( loException ) == "O", ;
													m.loException.Message, "Unknown error" )
			m.llSuccess = .F.
		Endtry
	    
		Return llSuccess  && Return operation result
	Endfunc


	*==============================================================
	* Method:    Cursor2Json
	* Purpose:   Export VFP cursor/table (or current work area) to JSON string
	* Params:    tcCursorName - (Optional) Cursor/table name, omit for current work area
	*            tnMaxRows    - Max rows to export (-1 = all)
	* Returns:   JSON string (array)
	*==============================================================
	Function Cursor2Json( tcCursorName As String , tnMaxRows As Integer) As String

	    Local lcJson		;
	    	, lnRows		;
	    	, lnRecno		;
	    	, lnFieldCount	;
	    	, laFields[1]	;
	    	, i  			;
	    	, loColl		;
	    	, loItem		;
	    	, lcFldName		;
	    	, lcFldType		;
	    	, lvValue		;
	    	, lcErrorMsg    

		m.lcErrorMsg	= ""
	    m.lcJson 		= ""
	    m.loColl 		= CreateObject("Collection")
	    
	    * Select work area
	    Local lcOldAlias	;
	    	, llAreaSwitched
	    
	    m.lcOldAlias 		= Alias()
	    m.llAreaSwitched	= .F.

	    If Vartype( m.tcCursorName ) = "C" And !Empty( m.tcCursorName )
	        If !Used( m.tcCursorName )
	            This.cLastError = "Cursor/table '" + tcCursorName + "' not open"
	            Return ""
	        Endif
	        Select ( m.tcCursorName )
	        m.llAreaSwitched = .T.
	    Endif

	    m.lnRows = 0
	    If Vartype( tnMaxRows ) != "N"
	        m.tnMaxRows = -1
	    Endif

	    If !Eof()
	        m.lnFieldCount = Fcount()
	        Dimension laFields[ lnFieldCount, 2 ]
	        For i = 1 To m.lnFieldCount
	            laFields[ i, 1 ] = Field( i )
	            laFields[ i, 2 ] = Type( laFields[ i, 1 ]) &&Vartype( Evaluate( Field( i ) ) )
	        Endfor

	        Scan
	            m.lnRows = m.lnRows + 1
	            If m.tnMaxRows > 0 And m.lnRows > m.tnMaxRows
	                Exit
	            Endif

	            m.loItem = CreateObject( "Empty" )

	            For i = 1 To m.lnFieldCount

	                m.lcFldName = laFields[ i, 1 ]
	                m.lcFldType = laFields[ i, 2 ]
	                m.lvValue   = Evaluate( lcFldName ) &&Field value
	                * Handle Null
	                If IsNull( m.lvValue )
	                    m.lvValue = Null
	                Else
	                    Do Case

	                    	* Trim string values
							Case m.lcFldType == "C"
								 m.lvValue = Alltrim( m.lvValue  )

							Case m.lcFldType == "D"
		                    * Date type ¡ú ISO format
	                             If Empty( m.lvValue ) then 
	                                m.lvValue = Null
	                             Endif 

							* DateTime type ¡ú ISO format
	                        Case m.lcFldType == "T" 
	                             If Empty( m.lvValue ) then 
	                                m.lvValue = Null
	                             Endif

							* Logical ¡ú keep .T./.F.
	                        Case m.lcFldType == "L"
	                            m.lvValue = Iif( m.lvValue, .T., .F. )

							* Memo detect JSON string ¡ú parse as object
	           				CASE m.lcFldType == "M" AND This.oNetJson.IsJsonString( m.lvValue , @lcErrorMsg )
								m.lvValue = This.Parse( m.lvValue )

							* Numeric types
	                        Case m.lcFldType $ "NIB"

							* Memo/binary fields
	                        Case m.lcFldType == "M" Or m.lcFldType == "G"

	                        Case m.lcFldType == "V"
	                            * Keep as character

	                        Otherwise
	                            * Convert other types to string
	                            m.lvValue = Transform( m.lvValue )
	                    Endcase
	                Endif

	                AddProperty( m.loItem, m.lcFldName, m.lvValue )
	            Endfor
	            loColl.Add( loItem )
	        Endscan
	    Endif

	    * Restore original work area
	    If m.llAreaSwitched And !Empty( m.lcOldAlias ) And Used( m.lcOldAlias )
	        Select ( m.lcOldAlias )
	    Endif
	    m.lcJson = This.Serialize( m.loColl )
	    
	    Return lcJson
	Endfunc

    *==============================================================
    * - Helper: IsValidVarName
    * - Purpose: Validate variable name
    *==============================================================
    Protected Function IsValidVarName( tcName As String) As Boolean
	    * - Rules: Cannot start with digit/cannot contain spaces
	    Return !Empty(tcName) And ;
	           Between(Len(tcName), 1, 128) And ;
	           Not IsDigit(Left(tcName, 1)) And ;
	           Not " " $ tcName
    Endfunc

	*==============================================================
	* Serialization core (performance critical)
	*==============================================================
	Protected Function SerializeValue(tuValue, loSB, tnLevel)
		Do Case
			Case VARTYPE(tuValue) $ "CM"
				This.EscapeString(tuValue, loSB)

			Case VARTYPE(tuValue) $ "NY"
				loSB.Add(Transform(tuValue))

			Case VARTYPE(tuValue) = "D"
				loSB.Add(This.FormatDate(tuValue))

			Case VARTYPE(tuValue) = "T"
				loSB.Add(This.FormatDateTime(tuValue))

			Case VARTYPE(tuValue) = "L"
				loSB.Add(IIF(tuValue, "true", "false"))

			Case VARTYPE(tuValue) = "O"
				This.HandleObject(tuValue, loSB, tnLevel)

			Case VARTYPE(tuValue) $ "X0U"
				loSB.Add("null")

			Case VARTYPE(tuValue) = "A"
				This.SerializeArray(tuValue, loSB, tnLevel)

			Otherwise
				This.CLastError = "Unsupported data type: " + VARTYPE(tuValue)
				Throw This.CLastError
		Endcase
	Endfunc

	*==============================================================
	* Unified object handling
	*==============================================================
	Protected Function HandleObject(toObject, loSB, tnLevel)
		* Check for circular references
		If This.IsReferenceExist(toObject)
			This.CLastError = "Circular reference detected: " + SYS(2015)
			Throw This.CLastError
		Endif
		
		* Add object reference
		This.AddReference(toObject)

		Try
			* Safely detect object type
			Local llIsCollection
			llIsCollection = .F.
			
			* Method 1: Check BaseClass property
			If TYPE("toObject.BaseClass") == "C"
				llIsCollection = (LOWER(toObject.BaseClass) == "collection")
			Endif
			
			* Method 2: Check for Count and Item methods (fallback)
			If !llIsCollection
				Try
					If TYPE("toObject.Count") == "N" AND TYPE("toObject.Item(1)") != "U"
						llIsCollection = .T.
					Endif
				Catch
					llIsCollection = .F.
				Endtry
			Endif

			* Handle collections vs regular objects
			If llIsCollection
				This.SerializeCollection(toObject, loSB, tnLevel)
			Else
				This.SerializeObject(toObject, loSB, tnLevel)
			Endif
		Catch To oException
			This.CLastError = "Object handling error: " + oException.Message
			If !EMPTY(oException.UserValue)
				This.CLastError = This.CLastError + " | " + oException.UserValue
			Endif
			Throw This.CLastError
		Finally
			* Ensure reference removal
			This.RemoveReference(toObject)
		Endtry
	Endfunc

	*==============================================================
	* Check if reference exists 
	*==============================================================
	Protected Function IsReferenceExist(toObject)
		Local i
		For i = 1 To This.nRefCount
			If This.HReferences[i] == toObject
				Return .T.
			Endif
		Endfor
		Return .F.
	Endfunc

	*==============================================================
	* Add object reference (dynamic expansion)
	*==============================================================
	Protected Function AddReference(toObject)
		This.nRefCount = This.nRefCount + 1
		
		* Dynamic array expansion
		If This.nRefCount > This.nRefCapacity
			This.nRefCapacity = This.nRefCapacity * 2
			Dimension This.HReferences[This.nRefCapacity]
		Endif
		
		This.HReferences[This.nRefCount] = toObject
	Endfunc

	*==============================================================
	* Remove object reference (efficient stack removal)
	*==============================================================
	Protected Function RemoveReference(toObject)
		If This.nRefCount > 0
			This.nRefCount = This.nRefCount - 1
			* No need to clear data, refcount will overwrite
		Endif
	Endfunc

	*==============================================================
	* Object serialization (optimized property traversal)
	*==============================================================
	Protected Function SerializeObject(toObject, loSB, tnLevel)
	    Local Array laProps[1]
	    Local lnCount, lcIndent, lcInnerIndent, i, lcPropName, lcValueType
	    Local llFirstProp, lnStartLen, lvValue, llSafeToRead

	    * Safely get object properties
	    Try
	        lnCount = AMEMBERS(laProps, toObject, 1)
	    Catch
	        lnCount = 0
	    Endtry
	    
	    If lnCount = 0
	        loSB.Add("{}")
	        Return
	    Endif

	    loSB.Add("{")
	    llFirstProp = .T.
	    lnStartLen  = loSB.Length()  && Record initial length

	    lcIndent = This.GetIndent(tnLevel)
	    lcInnerIndent = This.GetIndent(tnLevel + 1)

	    Local loTempSB
	    loTempSB = CREATEOBJECT("StringBuilder")
	    loTempSB.PrepareAdd(512)  && Pre-allocate

	    * Single pass property processing
	    For i = 1 To lnCount
	        If laProps[i, 2] = "Method"
	            Loop  && Skip methods
	        Endif

	        lcPropName = Lower( laProps[i, 1] )
	        
	        * --- Critical fix 1: Check property readability ---
	        llSafeToRead = PEMSTATUS(toObject, lcPropName, 5)  && 5=readable check
	        If Not llSafeToRead
	            Loop  && Skip unreadable properties
	        Endif

	        * Safely get property type
	        Try
	            lcValueType = TYPE("toObject." + lcPropName)
	        Catch
	            lcValueType = "U"
	        Endtry

	        * --- Critical fix 2: Unified exception handling ---
	        Try
	            lvValue = EVALUATE("toObject." + lcPropName)
	        Catch
	            lvValue = .Null.  && Use null on error
	        Endtry

	        * Batch process simple types
	        Do Case
		        Case lcValueType $ "CMNYDLT"
		            If Not llFirstProp
		                loTempSB.Add(",")
		                If This.LFormatted
		                    loTempSB.Add(CRLF + lcInnerIndent)
		                Endif
		            Else
		                llFirstProp = .F.
		            Endif

		            This.EscapeString(lcPropName, loTempSB)
		            loTempSB.Add(":")
		            This.SerializeValue(lvValue, loTempSB, tnLevel + 1)

		        Otherwise
		            * Complex types handle directly
		            If loTempSB.Length() > 0
		                loSB.Add(loTempSB.ToString())
		                loTempSB.Clear()
		                llFirstProp = .F.  && Properties already output
		            Endif
		            
		            If Not llFirstProp
		                loSB.Add(",")
		            Else
		                llFirstProp = .F.
		            Endif
		            
		            If This.LFormatted
		                loSB.Add(CRLF + lcInnerIndent)
		            Endif

		            This.EscapeString(lcPropName, loSB)
		            loSB.Add(":")
		            This.SerializeValue(lvValue, loSB, tnLevel + 1)
		        Endcase
	    Endfor

	    * Add remaining simple properties
	    If loTempSB.Length() > 0
	        loSB.Add(loTempSB.ToString())
	    Endif

	    If This.LFormatted And lnCount > 0
	        loSB.Add(CRLF + lcIndent)
	    Endif

	    loSB.Add("}")
	Endfunc

	*==============================================================
	* EscapeString optimization
	*==============================================================
	Protected Function EscapeString(tcString, loSB)
		If EMPTY(tcString)
            loSB.Add('""')
            Return
        Endif

        * Handle common escape characters first
        lcResult = tcString
        
        * Key optimization: Use STRTRAN for batch replacements
        lcResult = STRTRAN(lcResult, "\", "\\", 1, -1, 1)  && Backslash
        lcResult = STRTRAN(lcResult, '"', '\"', 1, -1, 1)  && Double quote
        lcResult = STRTRAN(lcResult, CHR(13) + CHR(10), "\n", 1, -1, 1)  && CRLF
        lcResult = STRTRAN(lcResult, CHR(10), "\n", 1, -1, 1)  && LF
        lcResult = STRTRAN(lcResult, CHR(13), "\n", 1, -1, 1)  && CR
        lcResult = STRTRAN(lcResult, CHR(9), "\t", 1, -1, 1)   && Tab
        lcResult = STRTRAN(lcResult, CHR(8), "\b", 1, -1, 1)   && Backspace
        lcResult = STRTRAN(lcResult, CHR(12), "\f", 1, -1, 1)  && Form feed

        * Handle other control chars (0-31)
        Local i
        For i = 0 To 31
            * Skip already handled special chars
            Do Case
                Case i = 8   && Backspace (handled)
                Case i = 9   && Tab (handled)
                Case i = 10  && LF (handled)
                Case i = 12  && FF (handled)
                Case i = 13  && CR (handled)
                Otherwise
                    * Replace with Unicode escape
                    lcResult = STRTRAN(lcResult, CHR(i), "\u" + PADL(Transform(i, "@0"), 4, "0"), 1, -1, 1)
            Endcase
        Endfor

        * Add quotes and output
        loSB.Add('"' + lcResult + '"')
    Endfunc
	
	*==============================================================
	* Array serialization (optimized for large arrays)
	*==============================================================
	Protected Function SerializeArray(taArray, loSB, tnLevel)
		Local lnDims, lnRows, lnCols, i, j

		* Safely get array dimensions
		Try
			lnDims = ALEN(taArray, 0)
			lnRows = ALEN(taArray, 1)
			lnCols = IIF(lnDims > 1, ALEN(taArray, 2), 0)
		Catch
			loSB.Add("[]")
			Return
		Endtry

		loSB.Add("[")

		Local lcIndent, lcInnerIndent
		lcIndent = This.GetIndent(tnLevel)
		lcInnerIndent = This.GetIndent(tnLevel + 1)

		* 1D array optimization - pre-calculate total length
		If lnDims = 1
			Local loTempSB, lnTotalSize
			loTempSB = CREATEOBJECT("StringBuilder")
			lnTotalSize = 0
			
			* Pre-calculate required memory
			For i = 1 To lnRows
				Try
					lnTotalSize = lnTotalSize + LEN(Transform(taArray[i]))
				Catch
				Endtry
			Endfor
			loTempSB.PrepareAdd(lnTotalSize + lnRows * 3)  && Add separator space

			For i = 1 To lnRows
				If i > 1
					loTempSB.Add(",")
					If This.LFormatted
						loTempSB.Add(CRLF + lcInnerIndent)
					Endif
				Endif

				* Safely handle array element
				Try
					* Batch process simple types
					Do Case
					Case TYPE("taArray[i]") $ "CMNYDLT"
						This.SerializeValue(taArray[i], loTempSB, tnLevel + 1)
					Otherwise
						loSB.Add(loTempSB.ToString())
						loTempSB.Clear()
						This.SerializeValue(taArray[i], loSB, tnLevel + 1)
					Endcase
				Catch
					loSB.Add(loTempSB.ToString())
					loTempSB.Clear()
					loSB.Add("null")
				Endtry
			Endfor

			If loTempSB.Length() > 0
				loSB.Add(loTempSB.ToString())
			Endif
		Else
			* 2D array handling - chunk processing
			Local lnBlockSize, iStart, iEnd
			lnBlockSize = 50  && 50 rows per block
			
			For i = 1 To lnRows
				If i > 1
					loSB.Add(",")
				Endif

				If This.LFormatted
					loSB.Add(CRLF + lcInnerIndent)
				Endif

				loSB.Add("[")

				* Column processing
				For j = 1 To lnCols
					If j > 1
						loSB.Add(",")
					Endif
					
					* Safely handle array element
					Try
						This.SerializeValue(taArray[i, j], loSB, tnLevel + 2)
					Catch
						loSB.Add("null")
					Endtry
				Endfor

				loSB.Add("]")
				
				* Flush buffer after each chunk
				If i % lnBlockSize = 0
					loSB.FlushBuffer()
				Endif
			Endfor
		Endif

		If This.LFormatted And (lnRows > 0 Or lnCols > 0)
			loSB.Add(CRLF + lcIndent)
		Endif

		loSB.Add("]")
	Endfunc

	*==============================================================
	* Collection serialization (enhanced error handling)
	*==============================================================
	Protected Function SerializeCollection(toCollection, loSB, tnLevel)
		Local lnCount, i

		* Safely get collection count
		Try
			lnCount = toCollection.Count
		Catch
			lnCount = 0
		Endtry

		loSB.Add("[")

		Local lcIndent, lcInnerIndent
		lcIndent = This.GetIndent(tnLevel)
		lcInnerIndent = This.GetIndent(tnLevel + 1)

		Local loTempSB
		loTempSB = CREATEOBJECT("StringBuilder")
		loTempSB.PrepareAdd(512)

		For i = 1 To lnCount
			Local luItem
			* Safely get collection item
			Try
				luItem = toCollection.Item(i)
			Catch
				luItem = .Null.
			Endtry

			If i > 1
				loTempSB.Add(",")
				If This.LFormatted
					loTempSB.Add(CRLF + lcInnerIndent)
				Endif
			Endif

			* Batch process simple types
			Try
				Do Case
				Case VARTYPE(luItem) $ "CMNYDLT"
					This.SerializeValue(luItem, loTempSB, tnLevel + 1)
				Otherwise
					loSB.Add(loTempSB.ToString())
					loTempSB.Clear()
					This.SerializeValue(luItem, loSB, tnLevel + 1)
				Endcase
			Catch
				loSB.Add(loTempSB.ToString())
				loTempSB.Clear()
				loSB.Add("null")
			Endtry
		Endfor

		If loTempSB.Length() > 0
			loSB.Add(loTempSB.ToString())
		Endif

		If This.LFormatted And lnCount > 0
			loSB.Add(CRLF + lcIndent)
		Endif

		loSB.Add("]")
	Endfunc

	*==============================================================
	* Helper: Generate indent string
	*==============================================================
	Protected Function GetIndent(tnLevel)
		If This.LFormatted
			Return Replicate( " ", tnLevel * This.NIndentSize )
		Endif
		Return ""
	Endfunc

	*==============================================================
	* Format date to ISO 8601 string
	*==============================================================
	Protected Function FormatDate( tdDate ) As String
		If Empty( m.tdDate ) Or Vartype( m.tdDate ) != "D"
			Return '"0000-00-00"'
		Endif
		Return '"' +  Transform( Year ( m.tdDate))		    + '-' + ;
				PADL( Transform( Month( m.tdDate)), 2, '0') + '-' + ;
				PADL( Transform( Day  ( m.tdDate)), 2, '0') + '"'
	Endfunc

	*==============================================================
	* Format datetime to ISO 8601 string
	*==============================================================
	Protected Function FormatDateTime(ttDateTime) As String
	
		If Empty( m.ttDateTime ) Or Vartype( m.ttDateTime ) != "T"
			Return '"0000-00-00T00:00:00"'
		Endif
		
		Return '"' +  Transform( Year 	( m.ttDateTime )) 			+ '-' + ;
				Padl( Transform( Month	( m.ttDateTime )), 2, '0') 	+ '-' + ;
				Padl( Transform( Day	( m.ttDateTime )), 2, '0') 	+ 'T' + ;
				Padl( Transform( Hour	( m.ttDateTime )), 2, '0') 	+ ':' + ;
				Padl( Transform( Minute	( m.ttDateTime )), 2, '0') 	+ ':' + ;
				Padl( Transform( Sec	( m.ttDateTime )), 2, '0') 	+ '"'
	Endfunc

	*==============================================================
	* - Method:    GetDecimalCount
	* - Purpose:   Get decimal places count in number
	* - Params:    tnValue - Numeric value
	* - Returns:   Decimal places count
	*==============================================================
	Function GetDecimalCount(tnValue As Number) As Integer
	    
		If Int( tnValue ) == tnValue
			Return 0  && Integer
		Endif
	    
		* Convert to string for processing
		Local lcValue As String
		m.lcValue = Transform(m.tnValue, "") && Remove thousand separators
		
		* Find decimal point position
		Local lnDotPos	As Integer;	
	    	, lnEPos 	As Integer
	    	
		m.lnDotPos = At(".", m.lcValue)
		m.lnEPos   = At("E", Upper(m.lcValue))
	    
		If m.lnDotPos > 0
			* Calculate decimal places
			If m.lnEPos > 0
				Return m.lnEPos - m.lnDotPos - 1
			Else
				Return Len( m.lcValue ) - m.lnDotPos
			Endif
		Endif
	    
		Return 0
	Endfunc

    *==============================================================
    * - Method:    GetLastError
    * - Purpose:   Get last error message
    * - Returns:   Error message string
    *==============================================================
    Function GetLastError() As String
             Return This.cLastError
    Endfunc

    *==============================================================
    * - Method:    Destroy
    * - Purpose:   Destroy object
    *==============================================================
    Function Destroy()
        If Vartype( This.oNetJson ) == "O"
			This.oNetJson.Dispose()
            This.oNetJson = .Null.
        Endif
        This.lInitialized = .F.
		
        * - Unload CLR runtime
        * - Check dependency files
        If !File( "clrhost.dll" )
            This.cLastError = "Dependency file 'clrhost.dll' not found"
            Return .Null.
        Endif
        
        Declare Integer ClrUnload In clrhost.dll    
        ClrUnload()
        Clear Dlls ClrUnload

    Endfunc

    *==============================================================
    * - Method:    CreateObjects
    * - Purpose:   Load DLL and create COM object
    * - Params:
    *       tcClassName - COM class name
    *       tcDllPath   - DLL file path
    * - Returns:   Object reference (success) or .NULL. (failure)
    *==============================================================
    Function Createobjects( tcClassName As String ;
                          , tcDllPath   As String )
            
        * - Prepare variables
        Local llSuccess ;
            , lcError   ;
            , lnSize    ;
            , loObject  ;
            , lnDispHandle
        
        m.llSuccess = .F.
        m.lcError   = ""
        
        * - Check dependency files
        If !File( "clrhost.dll" )
            This.cLastError = "Dependency file 'clrhost.dll' not found"
            Return .Null.
        Endif
        
        If !File( m.tcDllPath )
            This.cLastError = "Dependency file '" + m.tcDllPath + "' not found"
            Return .Null.
        Endif
        
        * - Declare API functions
        Declare Integer SetClrVersion In clrhost.dll String 
        Declare Integer ClrCreateInstanceFrom In clrhost.dll String, String, String@, Integer@
        
        * - Set CLR version
        SetClrVersion("v4.0.30319")  && Use .NET 4.0 runtime
        Clear Dlls SetClrVersion
        
        m.lcError = Space(1024)
        m.lnSize = 1024
        
        * - Create COM object instance
        m.lnDispHandle = ClrCreateInstanceFrom( Fullpath( m.tcDllPath ), ;
                                                m.tcClassName           , ;
                                                @lcError               , ;
                                                @lnSize )
        
        Clear Dlls ClrCreateInstanceFrom
        
        * - Handle creation result
        If m.lnDispHandle < 1 Then 
        
           m.lcError = Alltrim( Strconv(Strconv( m.lcError, 5), 6))
           m.lcError = Strtran( m.lcError, Chr(0), "")
            
            If Empty( m.lcError ) Then 
                Local Aerrs[1]
                Aerror( Aerrs )
                m.lcError = Aerrs[2]
            Endif
            
            This.cLastError = "Object creation failed " + tcClassName + "! Error:" + Chr(13) + lcError
            Return .Null.
            
        Else
            * - Convert to VFP object
            m.loObject = Sys(3096, lnDispHandle)
            Sys( 3097, m.loObject )  && Release COM handle
            Return m.loObject
        Endif
    Endfunc
    
Enddefine

*==============================================================
* String Builder class (reduce string concatenation overhead)
*==============================================================
Define Class StringBuilder As Custom
		Protected aParts[1], nCount, nTotalLength, nCapacity

		nCount = 0
		nTotalLength = 0
		nCapacity = 0
		Dimension aParts[16]

		* Pre-allocate memory
		Function PrepareAdd(tnSize)
				This.nCapacity = This.nCapacity + tnSize
		Endfunc

		Function Add(tcString)
				Local lnLen
				lnLen = LEN(tcString)

				If This.nCount >= ALEN(This.aParts)
					Local lnNewSize
					lnNewSize = ALEN(This.aParts) * 2
					Dimension This.aParts[lnNewSize]
				Endif

				This.nCount = This.nCount + 1
				This.aParts[This.nCount] = tcString
				This.nTotalLength = This.nTotalLength + lnLen
		Endfunc

		Function Clear
				This.nCount = 0
				This.nTotalLength = 0
				This.nCapacity = 0
		Endfunc

		Function Length
				Return This.nTotalLength
		Endfunc

		* Periodic memory buffer cleanup
		Function FlushBuffer
				* Not usually needed in VFP
				* Retained for compatibility
		Endfunc

		Function ToString
			    Local lcResult, i
			    If This.nCount = 0
			    	Return ""
			    Endif

			    lcResult = This.aParts[1]
			    For i = 2 To This.nCount
			    	lcResult = lcResult + This.aParts[i]
			    Endfor
			    Return lcResult
		Endfunc

Enddefine