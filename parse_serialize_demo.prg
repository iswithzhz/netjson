Clear

* ================== Initialization Setup ==================
Local lcPath
If _vfp.StartMode = 0
    * - Use project path in development environment
    m.lcPath = Justpath(_vfp.ActiveProject.Name)
Else
    * - Use executable path in runtime environment
    m.lcPath = Justpath( Sys(16, 0) )
Endif

Set Default To ( m.lcPath )               	  
Set Procedure To NetJson.prg && VIP user's additional JSON serializer file
Set Safety Off

* ================== Reuse NetJson Instance ==================
If Type( "oNetJson" ) != "O" Or Isnull( oNetJson ) Then 
    Public oNetJson 
    oNetJson = Createobject( "NetJson" )
Endif  

* ================== 1. Simple Object Test ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
{
  "name": "John Doe",
  "age": 30,
  "isStudent": false,
  "address": null
}
Endtext 

* ===== cJson -> oJson =======
    * - Measure time to verify efficiency
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson1 
    m.loJson1 = oNetJson.Parse( m.lcJson )
    
    * - Get error message on failure
    If Vartype( m.loJson1 ) != "O" Then 
       ? "1. Simple object deserialization failed        :" , oNetJson.clASTERROR
    Else 
       ? "1. Simple object deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif 

* ===== oJson -> cJson =======
    * - Measure time to verify efficiency
    Local lnHtTime 
    m.lnHtTime = Seconds()

    * - Serialize to JSON
    m.lcJson = oNetJson.Serialize( loJson1 )
    If Empty( oNetJson.CLastError )
        ? "1. Simple object serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        * - Save to file
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "1. SimpleObject.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "1. Simple object serialization failed    :", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif

* ================== 2. Nested Object Test ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
{
  "person": {
    "firstName": "Jane",
    "lastName": "Smith",
    "contact": {
      "email": "jane@example.com",
      "phone": "123-456-7890"
    }
  }
}
Endtext 

* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson2 
    m.loJson2 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson2 ) != "O" Then 
       ? "2. Nested object failed        :" , oNetJson.clASTERROR
    Else 
       ? "2. Nested object deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif 

* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson2 )
    If Empty( oNetJson.CLastError )
        ? "2. Nested object serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "2. NestedObject.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "2. Nested object serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif

* ================== 3. Array Structure Test ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
[
  "apple",
  "banana",
  "cherry",
  42,
  true,
  null
]
Endtext 

* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson3
    m.loJson3 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson3 ) != "O" Then 
       ? "3. Array structure failed      :" , oNetJson.clASTERROR
    Else 
       ? "3. Array structure deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif 

* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson3 )
    If Empty( oNetJson.CLastError )
        ? "3. Array structure serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "3. ArrayStructure.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "3. Array serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif

* ================== 4. Object Array Test ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
{
  "employees": [
    {"id": 1, "name": "Alice", "position": "Manager"},
    {"id": 2, "name": "Bob", "position": "Developer"},
    {"id": 3, "name": "Charlie", "position": "Designer"}
  ]
}
Endtext 

* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson4
    m.loJson4 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson4 ) != "O" Then 
       ? "4. Object array failed         :" ,oNetJson.clASTERROR
    Else 
       ? "4. Object array deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif

* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson4 )
    If Empty( oNetJson.CLastError )
        ? "4. Object array serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "4. ObjectArray.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "4. Object array serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif

* ================== 5. Escape Character Test ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
{
  "description": "This is a test of special characters: \nNew Line, \tTab, \\Backslash, \"Quote",
}
Endtext 

* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson5
    m.loJson5 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson5 ) != "O" Then 
       ? "5. Escape character test failed    :" , oNetJson.clASTERROR
    Else 
       ? "5. Escape character deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif

* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson5 )
    If Empty( oNetJson.CLastError )
        ? "5. Escape character serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "5. EscapeCharacters.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "5. Escape character serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif

* ================== 6. Complex String Combination Test ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
{
  "mixedContent": "Line 1\nLine 2\r\nLine 3\tTabbed\\Backslash\"Quoted\"",
  "specialChars": "~!@#$%^&*()_+`-=[]{}|;':,./<>?"
}
Endtext 

* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson6
    m.loJson6 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson6 ) != "O" Then 
       ? "6. Complex string combination failed       :" , oNetJson.clASTERROR
    Else 
       ? "6. Complex string deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif

* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson6 )
    If Empty( oNetJson.CLastError )
        ? "6. Complex string serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "6. ComplexString.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "6. Complex string serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif
    
* ================== 7. Null Value Test ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
{
  "emptyString": "",
  "nullValue": null,
  "emptyArray": [],
  "emptyObject": {}
}
Endtext 
* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson7
    m.loJson7 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson7 ) != "O" Then 
       ? "7. Null value test failed      :",oNetJson.clASTERROR
    Else 
       ? "7. Null value deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif

* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson7 )
    If Empty( oNetJson.CLastError )
        ? "7. Null value serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "7. NullValues.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "7. Null value serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif
    
* ================== 8. Extreme Value Test ==================
*3.4028235e38 |  3.40282350000001e38 - VFP uses double precision (8 bytes) ~15-17 significant digits
Local lcJson 
Text To m.lcJson Textmerge Noshow
{
  "maxInt": 2147483647,
  "minInt": -2147483648,
  "largeFloat": 3.4028235e38,
  "smallFloat": 1.17549435e-38,
  "zero": 0
}
Endtext 
* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson8
    m.loJson8 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson8 ) != "O" Then 
       ? "8. Extreme value test failed      :",oNetJson.clASTERROR
    Else 
       ? "8. Extreme value deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif
    
* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson8 )
    If Empty( oNetJson.CLastError )
        ? "8. Extreme value serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "8. ExtremeValues.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "8. Extreme value serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif

* ================== 9. Deep Nesting Structure Test ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
{
  "level1": {
    "level2": {
      "level3": {
        "level4": {
          "level5": {
          	"level6": {
          	  "level7": {
          		"level8": {
          		   "level9": {	
		            "value": "Deeply nested"
		          }
		        }
	          }
            }
          }
        }
      }
    }
  }
}
Endtext 
* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson9
    m.loJson9 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson9 ) != "O" Then 
       ? "9. Deep nesting failed         :" , oNetJson.clASTERROR
    Else 
       ? "9. Deep nesting deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif

* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson9 )
    If Empty( oNetJson.CLastError )
        ? "9. Deep nesting serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "9. DeepNesting.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "9. Deep nesting serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif
    
* ================== 10. Mixed Type Array Test ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
[
  "string",
  42,
  true,
  null,
  {"key": "value"},
  [1, 2, 3],
  3.14
]
Endtext 
* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson10
    m.loJson10 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson10 ) != "O" Then 
       ? "10. Mixed type array failed      :" , oNetJson.clASTERROR
    Else 
       ? "10. Mixed type array deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif
* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson10 )
    If Empty( oNetJson.CLastError )
        ? "10. Mixed type array serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "10. MixedTypeArray.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "10. Mixed type serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif

* ================== 11. Date/Time Test ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
{
  "date": "2023-12-31",
  "datetime": "2023-12-31T23:59:59"
}
Endtext 
* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson11
    m.loJson11 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson11 ) != "O" Then 
       ? "11. Date/time test failed      :" , oNetJson.clASTERROR
    Else 
       ? "11. Date/time deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif
* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson11 )
    If Empty( oNetJson.CLastError )
        ? "11. Date/time serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "11. DateTimeTest.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "11. Date/time serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif

* ================== 12. Comprehensive Test 1 ==================
Local lcJson 
Text To m.lcJson Textmerge Noshow
{
    "testCases": [
        { "name": "Number test", "value": 12345.6789, "expectedType": "number" },
        { "name": "Integer test", "value": 30, "expectedType": "integer" },
        { "name": "Boolean test", "value": true, "expectedType": "boolean" },
        { "name": "Null test", "value": null, "expectedType": "NULL" },
        { "name": "String test", "value": "Complex JSON example with multiple data types", "expectedType": "string" },
        { "name": "Special characters", "value": "Contains special chars: '\"\\/?<>|*", "expectedType": "string" },
        { "name": "Long string", "value": "Very long string...(254+ characters)", "expectedType": "long string" },
        { "name": "Date test", "value": "2023-11-15", "expectedType": "date" },
        { "name": "Datetime test", "value": "2023-11-15 15:34:46", "expectedType": "datetime" },
        { "name": "Object test", "value": { "name": "John", "age": 30 }, "expectedType": "object" },
        { "name": "Array test", "value": ["Java", "Python", "JavaScript"], "expectedType": "array" },
        { "name": "Nested object", "value": { "contact": { "phone": "13800138000" } }, "expectedType": "nested object" },
        { "name": "Mixed array", "value": [{ "project": "Enterprise system" }, "Other value"], "expectedType": "mixed array" }
    ]
}
Endtext 
* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()
    Local loJson12 
    m.loJson12 = oNetJson.Parse( m.lcJson )
    If Vartype( m.loJson12 ) != "O" Then 
       ? "12. Comprehensive test 1 failed      :" , oNetJson.clASTERROR
    Else 
       ? "12. Comprehensive test 1 deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif 

* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson12 )
    If Empty( oNetJson.CLastError )
        ? "12. Comprehensive test 1 serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "12. ComprehensiveTest1.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "12. Comprehensive test 1 serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif

* ================== 13. Comprehensive Test 2 ==================
Text To m.lcJson Noshow
{
    "basicInfo": {
        "name": "John",
        "age": 30,
        "gender": "Male",
        "isMarried": true,
        "birthDate": "1993-05-15",
        "idNumber": "110101199305151234",
        "contact": {
            "phone": "13800138000",
            "email": "john@example.com",
            "address": {
                "province": "Beijing",
                "city": "Beijing",
                "district": "Chaoyang",
                "detail": "88 Jianguo Road"
            }
        }
    },
    "workInfo": {
        "company": "ABC Tech",
        "position": "Senior Engineer",
        "joinDate": "2018-07-01",
        "department": "R&D",
        "salary": 25000,
        "projects": [
            {
                "name": "Enterprise System",
                "startDate": "2019-01-01",
                "endDate": "2020-06-30",
                "role": "Tech Lead",
                "techStack": ["Java", "Spring Boot", "MySQL"]
            },
            {
                "name": "E-commerce Platform",
                "startDate": "2020-07-01",
                "endDate": "2021-12-31",
                "role": "Architect",
                "techStack": ["Node.js", "Vue.js", "MongoDB"]
            }
        ]
    },
    "education": [
        {
            "school": "Peking University",
            "major": "Computer Science",
            "degree": "Bachelor",
            "startDate": "2010-09-01",
            "endDate": "2014-06-30"
        },
        {
            "school": "Tsinghua University",
            "major": "Software Engineering",
            "degree": "Master",
            "startDate": "2014-09-01",
            "endDate": "2017-06-30"
        }
    ],
    "skills": {
        "languages": ["Java", "Python", "JavaScript", "C++"],
        "databases": ["MySQL", "MongoDB", "Oracle"],
        "frameworks": ["Spring Boot", "Vue.js", "Django"],
        "certifications": [
            {
                "name": "PMP",
                "issuer": "PMI",
                "date": "2020-05-15"
            },
            {
                "name": "AWS Solutions Architect",
                "issuer": "Amazon Web Services",
                "date": "2021-08-20"
            }
        ]
    },
    "notes": "Complex JSON example for VFP conversion test",
    "specialCharTest": "Special characters: '\"\\/?<>|*",
    "nullTest": null,
    "numberTest": 12345.67,
    "boolTest": false,
    "dateTest": "2023-11-15",
    "datetimeTest": "2023-11-15 15:34:46",
    "floatPrecisionTest": 12345.9876543210
}
Endtext
* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson13 
    m.loJson13 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson13 ) != "O" Then 
       ? "13. Comprehensive test 2 failed:     ",oNetJson.clASTERROR
    Else 
       ? "13. Comprehensive test 2 deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
Endif 
* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson13 )
    If Empty( oNetJson.CLastError )
        ? "13. Comprehensive test 2 serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "13. ComprehensiveTest2.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "13. Comprehensive test 2 serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif

* ================== 14. Comprehensive Test 3 (Big Data) ==================
m.lcJson = Filetostr( "jsondata.json" )
* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson14 
    m.loJson14 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson14 ) != "O" Then 
       ? "14. Big data test failed      :", oNetJson.clASTERROR
    Else 
       ? "14. Big data deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif 
* ===== oJson -> cJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    m.lcJson = oNetJson.Serialize( loJson14 )
    If Empty( oNetJson.CLastError )
        ? "14. Big data serialization succeeded, time (sec) :", Seconds() - m.lnHtTime
        Local lcSavePath
        m.lcSavePath = Addbs( Getenv("TEMP") ) + "14. BigDataTest.json"
        StrToFile( m.lcJson, m.lcSavePath )
        Modify Command ( m.lcSavePath ) Nowait 
    Else
        ? "14. Big data serialization failed:", oNetJson.CLastError
        _cliptext = m.oNetJson.CLastError
    Endif
    

* ================== 15. jsondata4_error.json Parsing Error Test ==================
m.lcJson = Filetostr( "jsondata4_error.json" )

* ===== cJson -> oJson =======
    Local lnHtTime 
    m.lnHtTime = Seconds()

    Local loJson15 
    m.loJson15 = oNetJson.Parse( m.lcJson )

    If Vartype( m.loJson15 ) != "O" Then 
       ? "15. Big data test failed      :", oNetJson.clASTERROR
    Else 
       ? "15. Big data deserialization succeeded, time (sec) :", Seconds() - m.lnHtTime
    Endif 

On Shutdown Quit     

If _vfp.StartMode != 0  Then 
    Read Events
Endif