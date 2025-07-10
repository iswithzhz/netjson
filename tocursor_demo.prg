* ================== DEMO Instructions =================================================================
* 1. Initialization: Set paths based on runtime environment and include NetJson.prg procedure file. 
*    Create a global oNetJson object instance.
* 2. Example 1: Order Item Processing
*     - Parse a JSON string representing an order
*     - Iterate through each item in the order, adding two new properties: 
*       price (randomly generated) and itemTotal (price * quantity)
*     - Convert the items array to a cursor and display it
*     - Serialize the modified order object back to JSON
*     - Use JSONPath query ($.items[?(@.qty > 1)].sku) to find SKUs with quantity > 1
* 3. Example 2: Complex Order Processing
*     - Read JSON string from file "tocursor.json" and parse it
*     - Convert parsed orders array to cursor using two methods:
*         a. Single-row inference: Scan only first row for structure
*         b. Full scan: Scan all rows for structure (safer but slower)
*     - Serialize converted cursors back to JSON and save to temporary files
*     - Test nested collection conversion (order items) to cursor and serialization
* ================================================================================================= 

Clear 

* ================== Initialization Settings ==================
Local lcPath

*-- Set paths based on runtime environment
If _vfp.StartMode = 0
    * - Development environment: Use project path
    m.lcPath = JustPath(_vfp.ActiveProject.Name)
Else
    * - Runtime environment: Use executable path
    m.lcPath = JustPath(Sys(16, 0))
EndIf

*-- Set default path and include files
Set Default To (m.lcPath)
Set Procedure To NetJson.prg Additive  && Additional JSON serialization tool for VIP users
Set Safety Off

* ================== Reuse NetJson Instance ==================
If Type("oNetJson") != "O" Or IsNull(oNetJson)
    Public oNetJson 
    oNetJson = CreateObject( "NetJson" )
EndIf 

* - Example 1: Order Item Processing (cjson->ojson->add nodes->ToCursor->cjson loop)
* ====================================================== 
Local lcJson
m.lcJson = '{"orderId":"ORD-2023","items":[{"sku":"P123","qty":2},{"sku":"P456","qty":1}],"total":99.99}'
oOrder = oNetJson.Parse(m.lcJson)

* - Add price/itemTotal Nodes
For m.lni = 1 To oOrder.items.Count
    * Calculate total using random price and current item quantity
    *AddProperty( oOrder.items[m.lni], "price"		, Int( Rand() * 100 ) )	&&0 decimals      I
    AddProperty( oOrder.items[m.lni], "price"		, Rand() * 100 ) 		&&2 decimals      N
    *AddProperty( oOrder.items[m.lni], "price"		, Rand() * 100 * 0.12 ) &&4 decimals      Y
    *AddProperty( oOrder.items[m.lni], "price"		, Rand() * 100 * 0.12 * 0.34 ) &&6 decimals B
    AddProperty( oOrder.items[m.lni], "itemTotal"	, oOrder.items[m.lni].price * oOrder.items[m.lni].qty )
EndFor 
?"===================== Added Node Values =========================="
?"price: " 		, oOrder.items[1].price
?"itemTotal: "	, oOrder.items[1].itemTotal

* - oJson -> cursor
If oNetJson.ToCursor( oOrder.items, "ORDER_ITEMS" ) Then 
	Browse Name ORDER_ITEMS Nowait 
Else 
    ?"ToCursor failed: " , oNetJson.cLastError
    Set Step On 
Endif 
?"===================== oJson -> cJson =========================="
* - oJson -> cJson
m.lcJson = oNetJson.Serialize( m.oOrder )
If Empty( m.lcJson ) Then
   ?"Serialization failed: " , oNetJson.cLastError
   Set Step On 
Else 
	? "Modified Order Json:"
	? m.lcJson
Endif 

?"===================== Json Path Query =========================="
Local lcaResult;
	, lcJsonPath

m.lcJsonPath = "$.items[?(@.qty > 1)].sku"
m.lcaResult  = oNetJson.Query( m.lcJson , m.lcJsonPath )
If Empty( m.lcaResult ) Then 
   ? "Query found no results: " , oNetJson.cLastError
   Set Step On 
Else 
	Local loQuery
	m.loQuery = oNetJson.Parse( m.lcaResult )
	? "Query expression: " + m.lcJsonPath
	For m.lni = 1 To m.loQuery.Count
	    ? "Matching SKU: " , Transform(m.lni) , m.loQuery[m.lni]
	EndFor 
	?
Endif 

?"================== Example 2: Complex Order Processing =========================="

* --------- cJson-> oJson ---------
m.lcJson = Filetostr( "tocursor.json" )
* - Parse JSON
oData = oNetJson.Parse(m.lcJson)
If Isnull(oData)
    ?"Parse failed: ", oNetJson.cLastError 
     Set Step On 
Endif
* ============= oJson -> Cursor - Single-row Inference =============
If oNetJson.ToCursor( oData.orders, "Orders" ) Then 
    Select orders
    Browse Normal Title "ToCursor Conversion Result (Single-row Inference)" Nowait
Else
   ?"ToCursor failed: " , oNetJson.cLastError 
    Set Step On 
Endif
* --------- Cursor -> cJson (Cursor2Json.json) ---------
m.lcFjson = Addbs( Getenv("TEMP") ) + "SingleRowInference_Cursor2Json.json"
m.lcJson  = oNetJson.Cursor2Json( "Orders" ) 
StrtoFile( m.lcJson ,  m.lcFjson )
Modify Command ( m.lcFjson ) Nowait 


* ============= oJson -> Cursor - Full Scan ============= 
If oNetJson.ToCursor(oData.orders, "Orders2", -1)
    Select Orders2
    Browse Normal Title "ToCursor Conversion Result (Full Scan)" Nowait
    Display Structure
Else
    ?"ToCursor failed: " , oNetJson.cLastError
     Set Step On 
Endif
* --------- oJson -> cJson (Serialize.json) ---------
m.lcFjson = Addbs( Getenv("TEMP") ) + "FullScan_Serialize.json"
m.lcJson  = oNetJson.Serialize( oData )
StrtoFile( m.lcJson ,  m.lcFjson )
Modify Command ( m.lcFjson ) Nowait 

*  ============= Test Nested Collection Conversion  =============
If oNetJson.ToCursor( oData.orders.Item(1).items, "OrderItems")
    Select OrderItems
    Browse Normal Title "Nested Items Conversion Result" Nowait
Else
    ?"Nested collection conversion failed: " , oNetJson.cLastError
     Set Step On 
Endif

* --------- Cursor -> cJson (OrderItems.json) ---------
m.lcFjson = Addbs( Getenv("TEMP") ) + "OrderItems1_Cursor2Json.json"
m.lcJson  = oNetJson.Cursor2Json( "OrderItems" ) 
StrtoFile( m.lcJson , m.lcFjson )
Modify Command ( m.lcFjson ) Nowait 

* --------- oJson -> cJson (OrderItems2.json) ---------
m.lcFjson = Addbs( Getenv("TEMP") ) + "OrderItems1_Serialize.json"
m.lcJson  = oNetJson.Serialize( oData.orders.Item(1).items )
StrtoFile( m.lcJson , m.lcFjson )
Modify Command ( m.lcFjson ) Nowait