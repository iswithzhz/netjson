* ================== Initialization Settings ==================
Local lcPath
If _vfp.StartMode = 0
    * Use project path in development environment
    m.lcPath = Justpath(_vfp.ActiveProject.Name)
Else
    * Use executable path in runtime environment
    m.lcPath = Justpath(Sys(16, 0))
Endif

Set Default To ( m.lcPath )
Set Procedure To NetJson.prg  && Include JSON processing library

* ================== Reuse NetJson Instance ==================
If Type("oNetJson") != "O" Or Isnull(oNetJson)
    Public oNetJson
    oNetJson = Createobject("NetJson")
Endif 

Clear 

* ================== Basic Query Demo ==================
Local lcJson, lcJsonPath
m.lcJson = '{"orderId":"ORD-2023","items":[{"sku":"P123","qty":2},{"sku":"P456","qty":1}],"total":99.99}'
m.lcJsonPath = "$.items[?(@.qty > 1)].sku"

Local lcaResult
m.lcaResult = oNetJson.Query( m.lcJson , m.lcJsonPath)

If !Empty( m.lcaResult ) Then 
    Local loJson
    m.loJson = oNetJson.Parse( m.lcaResult )

    ? "===== Basic Query Demo ====="
    ? "Query expression: " + m.lcJsonPath
    For m.lni = 1 To m.loJson.Count
        ? "SKU " + Transform(m.lni) + ":", m.loJson[m.lni]
    Endfor 
    ?
Else 
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif 

* ================== Complex JSON Structure ==================
m.lcJson = Filetostr( "Query_Demo.json" )

* ================== Query Examples ==================
? "===== JSONPath Query Examples ====="
?

* Example 1: Get all company names
m.lcaResult = oNetJson.Query(m.lcJson, "$[*].company")
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "1. All company names:"
    For m.lni = 1 To m.loJson.Count
        ? "  Company " + Transform(m.lni) + ":", m.loJson[m.lni]
    Endfor 
    ?
Else 
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif

* Example 2: Get department manager email
* Using single quotes to wrap JSONPath, double quotes inside
m.lcJsonPath = '$[0].departments[?(@.name=="Sales")].manager.contact.email'
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)

If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "2. Department manager email:"
    For m.lni = 1 To m.loJson.Count
        ? "  Email:", m.loJson[m.lni]
    Endfor 
    ?
Else 
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif

* Example 3: Get special character field
* Using single quotes to wrap JSONPath
m.lcJsonPath = '$[0]["special/chars"]'
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "3. Special character field:"
    For m.lni = 1 To m.loJson.Count
        ? "  Content:", m.loJson[m.lni]
    Endfor 
    ?
Else 
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif

* Example 4: Recursively get all tags
m.lcaResult = oNetJson.Query(m.lcJson, "$..tags[*]")
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "4. All tags:"
    For m.lni = 1 To m.loJson.Count
        ? "  Tag " + Transform(m.lni) + ":", m.loJson[m.lni]
    Endfor 
    ?
Else 
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif 

* Example 5: Conditional query - Companies founded before 2000
m.lcaResult = oNetJson.Query(m.lcJson, "$[?(@.founded < 2000)].company")
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "5. Companies founded before 2000:"
    For m.lni = 1 To m.loJson.Count
        ? "  Company:", m.loJson[m.lni]
    Endfor 
    ?
Else 
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif

* Example 6: Nested array query - Departments with >20 employees
m.lcaResult = oNetJson.Query(m.lcJson, "$..departments[?(@.employees > 20)].name")

If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse( m.lcaResult )
    ? "6. Large departments (employees >20):"
    For m.lni = 1 To m.loJson.Count
        ? "  Department:", m.loJson[m.lni]
    Endfor 
    ?
Else 
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif

* ================== Fixed Query Examples ==================
? "===== JSONPath Query Examples (Fixed) ====="
?

* Example 7: Companies founded before 2000 (non-parameterized)
m.lcJsonPath = '$[?(@.founded < 2000)].company'
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "7. Companies founded before 2000:"
    For m.lni = 1 To m.loJson.Count
        ? "  Company:", m.loJson[m.lni]
    Endfor 
    ?
Else 
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif

* Example 8: Departments with 10-20 employees (Chr(38) = &)
m.lcJsonPath = '$..departments[?(@.employees >= 10 '+ Chr(38) + Chr(38) +' @.employees <= 20)].name'
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "8. Medium departments (10-20 employees):"
    For m.lni = 1 To m.loJson.Count
        ? "  Department:", m.loJson[m.lni]
    Endfor 
    ?
Else
    ? "8. No departments meet criteria"
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif 

* Example 9: Find Sales department ID
* Recursive search across all companies
m.lcJsonPath = "$..departments[?(@.name == 'Sales')].id"
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "9. Department ID query (Sales):"
    For m.lni = 1 To m.loJson.Count
        ? "  Department ID:", m.loJson[m.lni]
    Endfor 
    ?
Else
    ? "No departments found"
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif

* Example 10: High-salary positions (>150K Regional Director)
m.lcJsonPath = "$..employees[?(@.position == 'Regional Director' "+ Chr(38) + Chr(38) + "@.salary > 150000)]"
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "10. High-salary employees (>150K Regional Director):"
    
    * Correctly handle object arrays
    For m.lni = 1 To m.loJson.Count
        ? "  Employee ID:", m.loJson[m.lni].id
        ? "  Name:", m.loJson[m.lni].name
        ? "  Position:", m.loJson[m.lni].position
        ? "  Salary:", TRANSFORM(m.loJson[m.lni].salary)
        ? "  -------------------"
    Endfor 
    ?
Else
    ? "10. No matching employees found"
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif 

* Example 11: High-salary employees (DEPT-001 & salary>150K) (using 'and')
* Fix: Correctly handle returned object arrays
m.lcJsonPath = "$..employees[?(@.department == 'DEPT-001'" + Chr(38) + Chr(38) + '@.salary > 150000)]'
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "11. High-salary employees (DEPT-001 & salary>150K):"

    * Process object results - iterate directly
    For m.lni = 1 To m.loJson.Count
        ? "  Employee ID:", m.loJson[m.lni].id
        ? "  Name:", m.loJson[m.lni].name
        ? "  Position:", m.loJson[m.lni].position
        ? "  Salary:", TRANSFORM(m.loJson[m.lni].salary)
        ? "  -------------------"
    Endfor 
    ?
Else
    ? "11. No matching employees found"
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif 

* Example 12: Get all CEO compensation plans
m.lcJsonPath = '$..ceo.compensation'
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "12. CEO compensation plans:"
    For m.lni = 1 To m.loJson.Count
        ? "  CEO " + Transform(m.lni) + ":"
        ? "    Salary:", TRANSFORM(m.loJson[m.lni].salary)
        ? "    Bonus:", TRANSFORM(m.loJson[m.lni].bonus)
        ? "    Stock Options:", TRANSFORM(m.loJson[m.lni].stockOptions)
    Endfor 
    ?
Else 
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif  

* Example 13: Find R&D department projects
* Fix R&D ampersand handling
m.lcJsonPath = "$..departments[?(@.name == 'R&D')].projects[*].name"
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "13. R&D department projects:"
    For m.lni = 1 To m.loJson.Count
        ? "  Project:", m.loJson[m.lni]
    Endfor 
    ?
Else 
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif  

* Example 14: Find active company acquisitions
* Fix: Add array index
m.lcJsonPath = '$[?(@.active == true)]..acquisitions[*].company'
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "14. Active company acquisitions:"
    For m.lni = 1 To m.loJson.Count
        ? "  Acquired company:", m.loJson[m.lni]
    Endfor 
    ?
Else
    ? "14. No acquisitions found"
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif 

* Example 15: Find employees with Machine Learning skills
* Fix: Use correct array query syntax
m.lcJsonPath = "$..employees[?(@.skills[?(@ == 'Machine Learning')])].name"
m.lcaResult  = oNetJson.Query( m.lcJson, m.lcJsonPath )
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse( m.lcaResult )
    ? "15. Employees with Machine Learning skills:"
    For m.lni = 1 To m.loJson.Count
        ? "  Employee:", m.loJson[m.lni]
    Endfor 
    ?
Else
    ? "15. No employees with Machine Learning skills"
    ?"Failure reason: " +  oNetJson.clASTERROR
    _cliptext = oNetJson.clASTERROR
Endif  

* Example 16: Find last updated metadata
m.lcJsonPath = '$..metadata.last_updated'
m.lcaResult  = oNetJson.Query( m.lcJson, m.lcJsonPath )
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "16. Last updated time:"
    For m.lni = 1 To m.loJson.Count
        ? "  Update time:", m.loJson[m.lni]
    Endfor 
    ?
Else
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif

* Example 17: Find Boston lab equipment
* Fix: Use correct recursive search syntax
m.lcJsonPath = "$..labs[?(@.location == 'Boston')].equipment[*]"
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "17. Boston lab equipment:"
    For m.lni = 1 To m.loJson.Count
        ? "  Equipment:", m.loJson[m.lni]
    Endfor 
    ?
Else
    ? "17. No equipment found in Boston lab"
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif 

* Example 18: Find Quantum Ventures investment amounts
* Fix: Use correct array search syntax
m.lcJsonPath = "$..fundingRounds[?(@.investors[?(@ == 'Quantum Ventures')])].amount"
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "18. Quantum Ventures investments:"
    For m.lni = 1 To m.loJson.Count
        ? "  Amount:", TRANSFORM(m.loJson[m.lni])
    Endfor 
    ?
Else
    ? "18. No investments found"
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif 

* Example 19: High-salary employees (combined criteria)
* Reuse Example 11 query
m.lcJsonPath = "$..employees[?(@.department == 'DEPT-001' " + Chr(38) + Chr(38) + " @.salary > 150000)]"
m.lcaResult = oNetJson.Query(m.lcJson, m.lcJsonPath)
If !Empty( m.lcaResult ) Then 
    m.loJson = oNetJson.Parse(m.lcaResult)
    ? "19. High-salary employees (DEPT-001 & salary>150K):"
    For m.lni = 1 To m.loJson.Count
        ? "  Employee:", m.loJson[m.lni].name, "Salary:", TRANSFORM(m.loJson[m.lni].salary)
    Endfor 
    ?
Else
    ? "19. No high-salary employees found"
    ?"Failure reason: " +  oNetJson.clASTERROR
Endif 

* ================== Completion ==================
? 
? "All query demos completed!"