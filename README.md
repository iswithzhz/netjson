*==============================================================*

* Class:        NetJson 
* Parent:       Custom
* Dependencies: NetJson.dll (C# COM Component) | ClrHost.dll(7.29)
* Environment:  Net4.61 | VFP9 7423
* Author:       ZHZ
* Description:  Uses C# component to parse JSON and create VFP objects
*               cJson → oJson | oJson → cJson | oJson → Cursor | Cursor → cJson

* cJson → oJson Type Mapping:
	* Node        → VFP Empty class          O
	* Array       → VFP Collection class     O
	* Date        → VFP date literal {^...}  D
 	* DateTime    → VFP date literal {^...}  T
	* Null        → Null                     L
	* Boolean     → .T./.F.                  L 
    * Numeric     → 123.45                   N
    * HighPrecision → 0.132456789            N
  
* oJson → Cursor Type Mapping:
	* String      → Determine C(254) or M
	* Date/Time   → Maintain original type (D/T) 
	* Object      → Convert to M type (store JSON)
	* Numeric     → Select I/N/B type
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
* Cursor → oJson → cJson
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
* 1. Precise JSON validation and error location
	NetJson.dll strictly validates JSON against standards;
	Detects syntax errors (mismatched brackets, unquoted fields, extra commas);
	Pinpoints error locations for efficient debugging;
	Automated validation saves time with complex JSON structures.
	
* 2. High-performance parsing with lower latency
	Optimized algorithms for large JSON data and high-frequency scenarios;
	Reduces parsing delays in real-time systems (API responses, data imports);
	Critical for WsServer.ProcessRequest() JSON handling.
	
* 3. Native VFP object support for full OOP access
	Converts JSON directly to VFP object model;
	Enables intuitive property access (e.g., obj.name instead of manual parsing);
	Supports nested structures via Collection classes;
	Simplifies code logic with object-oriented syntax.
	
* 4. Strict type mapping for data safety
	Ensures accurate VFP type conversion (JSON true → .T., numbers → numeric);
	Prevents runtime errors from implicit type conversions;
	Improves code maintainability with explicit typing.
*==============================================================
