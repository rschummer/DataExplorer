* <summary>
*	Abstract Data Management class.  Must inherit
*	from this for any Data providers you have.
* </summary>
*
* RAS 3-Jul-2011, made pass for m.Variables
*
#include "DataExplorer.h"
#include "foxpro.h"

DEFINE CLASS DatabaseMgmt AS Session
	ServerName        = SPACE(0)
	DatabaseName      = SPACE(0)
	UserName          = SPACE(0)
	UserPassword      = SPACE(0)
	ConnectionString  = SPACE(0)
	
	SortObjects       = .T.
   
   * RAS 29-Jul-2006, added new property to allow columns to sort separate from
   * the other database objects (tables, views, stored procedures and functions).
   SortColumns       = .F.
   
	Verbose           = .T.

	IgnoreErrors      = .F.

	LastError         = SPACE(0)
	QueryResultOutput = SPACE(0)
   
   * RAS 3-Jul-2011, added new property to determine if the BROWSE grid is read-only or not.
   * This allows developers extending the BROWSE feature more control. For instance, developers
   * might want it read/write, but if they include this in a released app it might be used
   * by end-users in read-only configuration.
   lBrowseReadOnly   = .F.
	
	
	PROCEDURE Destroy()
		THIS.Disconnect()
	ENDPROC
	
	FUNCTION Connect(cServer, cDatabase, cUserName, cPassword) AS Boolean
	ENDFUNC
	
	FUNCTION Disconnect()
	ENDFUNC

	FUNCTION ClearErrors()
		THIS.LastError = ''
	ENDFUNC
	
	* strip password for a connection string
	FUNCTION StripPassword(cStr as String)
		LOCAL cNewValue
		LOCAL i
		LOCAL cVal
		
		cNewValue = SPACE(0)
      
		FOR i = 1 TO GETWORDCOUNT(m.cStr, ';')
			cVal = GETWORDNUM(m.cStr, i, ';')
			
         IF NOT (UPPER(LEFT(m.cVal, 8)) = "PASSWORD" OR UPPER(LEFT(m.cVal, 3)) = "PWD")
				cNewValue = m.cNewValue + IIF(EMPTY(m.cNewValue), '', ';') + GETWORDNUM(m.cStr, i, ';')
			ELSE
				IF AT('=', m.cVal) > 0
					cNewValue = m.cNewValue + IIF(EMPTY(m.cNewValue), '', ';') + LEFT(m.cVal, AT('=', m.cVal)) + "***"
				ENDIF
			ENDIF
		ENDFOR
		
		RETURN m.cNewValue
	ENDFUNC
	
	FUNCTION SetError(cErrorMsg)
		LOCAL i
		LOCAL cErrorMsg
		LOCAL nErrorCnt
		LOCAL ARRAY aErrorList[1]
		
		IF VARTYPE(m.cErrorMsg) <> 'C' OR EMPTY(m.cErrorMsg)
			cErrorMsg = SPACE(0)
			nErrorCnt = AERROR(aErrorList)
			
         FOR i = 1 TO m.nErrorCnt
				IF aErrorList[i, 1] == 1526 && ODBC error
					IF ISNULL(aErrorList[i, 3])
						cErrorMsg = m.cErrorMsg + IIF(EMPTY(m.cErrorMsg), SPACE(0), CHR(10)) + "Problem executing query"
					ELSE
						cErrorMsg = m.cErrorMsg + IIF(EMPTY(m.cErrorMsg), SPACE(0), CHR(10)) + SUBSTR(aErrorList[i, 3], RAT(']', aErrorList[i, 3]) + 1)
					ENDIF
				ELSE
					cErrorMsg = m.cErrorMsg + IIF(EMPTY(m.cErrorMsg), '', CHR(10)) + aErrorList[i, 2]
				ENDIF
			ENDFOR
		ENDIF

		THIS.LastError = m.cErrorMsg
	ENDFUNC


	FUNCTION GetAvailableServers()
		LOCAL oObjCollection
		LOCAL oException AS Exception
		
      oObjCollection = CREATEOBJECT("ServerCollection")

		TRY
			IF !THIS.OnGetAvailableServers(m.oObjCollection)
				oObjCollection = .NULL.
			ENDIF
         
		CATCH TO m.oException
			oObjCollection = .NULL.
			
         * ignore error
			IF !THIS.IgnoreErrors
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			ENDIF
		ENDTRY
	
		RETURN m.oObjCollection
	ENDFUNC

	FUNCTION GetDatabases(cServerName)
		LOCAL oObjCollection

		oObjCollection = CREATEOBJECT("DatabaseCollection")
		THIS.OnGetDatabases(m.oObjCollection)
		
		RETURN m.oObjCollection
	ENDFUNC

	FUNCTION GetTables()
		LOCAL oObjCollection
		LOCAL oException AS Exception
		
      oObjCollection = CREATEOBJECT("TableCollection")

		TRY
			IF !THIS.OnGetTables(m.oObjCollection)
				oObjCollection = .NULL.
			ENDIF
		
      CATCH TO oException
			oObjCollection = .NULL.
			
         IF !THIS.IgnoreErrors
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			ENDIF
		ENDTRY
		
		RETURN m.oObjCollection
	ENDFUNC

	FUNCTION GetViews()
		LOCAL oObjCollection
		LOCAL oException AS Exception
		
      oObjCollection = CREATEOBJECT("ViewCollection")
		
		TRY
			IF !THIS.OnGetViews(m.oObjCollection)
				oObjectCollection = .NULL.
			ENDIF
		
      CATCH TO m.oException
			IF !THIS.IgnoreErrors
				oObjectCollection = .NULL.
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			ENDIF
		ENDTRY
		
		RETURN m.oObjCollection
	ENDFUNC

	FUNCTION GetStoredProcedures()
		LOCAL oObjCollection
		LOCAL oException AS Exception
		
      oObjCollection = CREATEOBJECT("StoredProcCollection")

		TRY
			IF !THIS.OnGetStoredProcedures(m.oObjCollection)
				oObjCollection = .NULL.
			ENDIF
		
      CATCH TO oException
			oObjCollection = .NULL.
			
         IF NOT THIS.IgnoreErrors
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			ENDIF
		ENDTRY
		
		RETURN m.oObjCollection
	ENDFUNC

	FUNCTION GetParameters(cStoredProcName, cOwner)
		LOCAL oObjCollection
		LOCAL oException AS Exception
		
      oObjCollection = CREATEOBJECT("ParameterCollection")
		
		TRY
			IF !THIS.OnGetParameters(m.oObjCollection, m.cStoredProcName, m.cOwner)
				oObjCollection = .NULL.
			ENDIF
		
      CATCH TO m.oException
			oObjCollection = .NULL.
			
         IF !THIS.IgnoreErrors
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			ENDIF
		ENDTRY
		
		RETURN m.oObjCollection
	ENDFUNC

	FUNCTION GetFunctions()
		LOCAL oObjCollection
		LOCAL oException AS Exception
		
      oObjCollection = CREATEOBJECT("FunctionCollection")

		TRY
			IF NOT THIS.OnGetFunctions(m.oObjCollection)
				oObjCollection = .NULL.
			ENDIF
		
      CATCH TO oException
			oObjCollection = .NULL.
         
			IF !THIS.IgnoreErrors
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			ENDIF
		ENDTRY
		
		RETURN m.oObjCollection
	ENDFUNC

	FUNCTION GetFunctionParameters(cFuncName, cOwner)
		LOCAL oObjCollection
		LOCAL oException AS Exception
		
      oObjCollection = CREATEOBJECT("ParameterCollection")
		
      TRY
			IF !THIS.OnGetFunctionParameters(m.oObjCollection, m.cFuncName, m.cOwner)
				oObjCollection = .NULL.
			ENDIF			
		
      CATCH TO m.oException
			oObjCollection = .NULL.
			
         IF !THIS.IgnoreErrors
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			ENDIF
		ENDTRY
		
		RETURN m.oObjCollection
	ENDFUNC

	FUNCTION GetSchema(cTableName, cOwner)
		LOCAL oObjCollection
		LOCAL oException AS Exception
		
      oObjCollection = CREATEOBJECT("ColumnCollection")
		
      TRY
			IF !THIS.OnGetSchema(m.oObjCollection, m.cTableName, m.cOwner)
				oObjCollection = .NULL.
			ENDIF

		CATCH TO oException
			oObjCollection = .NULL.
		
      	IF !THIS.IgnoreErrors
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			ENDIF
		ENDTRY
			
		
		RETURN m.oObjCollection
	ENDFUNC
	
	FUNCTION GetStoredProcedureDefinition(cStoredProcName, cOwner) AS String
		RETURN NVL(THIS.OnGetStoredProcedureDefinition(m.cStoredProcName, m.cOwner), '')
	ENDFUNC

	FUNCTION GetFunctionDefinition(cFunctionName, cOwner) AS String
		RETURN NVL(THIS.OnGetFunctionDefinition(m.cFunctionName, m.cOwner), '')
	ENDFUNC

	FUNCTION GetViewDefinition(cViewName, cOwner) AS String
		RETURN NVL(THIS.OnGetViewDefinition(m.cViewName, m.cOwner), '')
	ENDFUNC
	
	FUNCTION BrowseData(cTableName, cOwner)
		IF THIS.OnBrowseData(m.cTableName, m.cOwner)
         this.OnBrowseForm(m.cTableName, m.cOwner)
      ENDIF 
	ENDFUNC

	FUNCTION DesignTable(cTableName, cOwner)
		THIS.OnDesignTable(m.cTableName, m.cOwner)
	ENDFUNC

	FUNCTION NewTable()
		THIS.OnNewTable()
	ENDFUNC

	FUNCTION DesignView(cViewName, cOwner)
		THIS.OnDesignView(m.cViewName, m.cOwner)
	ENDFUNC

	FUNCTION NewView()
		THIS.OnNewView()
	ENDFUNC

	FUNCTION RunStoredProcedure(m.cStoredProcName, m.cOwner)
		LOCAL oParamList
		LOCAL lSuccess

		lSuccess = .T.
	
		oParamList = THIS.GetParameters(m.cStoredProcName, m.cOwner)

		IF VARTYPE(m.oParamList) == 'O' AND m.oParamList.Count > 0
			DO FORM StoredProcParameters WITH m.oParamList TO m.lSuccess
		ENDIF
		
      IF m.lSuccess
			THIS.OnRunStoredProcedure(m.cStoredProcName, m.cOwner, m.oParamList)
		ENDIF
	ENDFUNC
	
	* ExecuteQuery always runs the query in the current
	* datasession.  Primarily for internal queries -- for example,
	* to retrieve schema information.
*!*		PROTECTED FUNCTION ExecuteQuery(cSQL, cAlias)
*!*			RETURN THIS.RunQuery(cSQL, cAlias, THIS.DataSessionID)
*!*		ENDPROC

	PROCEDURE RunQuery(cSQL)
		THIS.OnRunQuery(m.cSQL)
	ENDPROC
	
	PROCEDURE OnRunQuery(cSQL)
		DO FORM RunQuery WITH THIS, m.cSQL
	ENDPROC
	
	* Use RunQuery to execute a query that returns
	* a result in a specified dataset. 
	FUNCTION ExecuteQuery(cSQL, cAlias, nDataSessionID)
		LOCAL nSelect
		LOCAL nSaveDataSessionID
		LOCAL lSuccess
		LOCAL nErrorCnt
		LOCAL oResultCollection
		LOCAL aErrorList[1]
		
	
		nSelect = SELECT()
		
		IF VARTYPE(m.cAlias) <> 'C' OR EMPTY(m.cAlias)
			cAlias = "SQLResults"
		ENDIF
		
		cAlias             = CHRTRAN(m.cAlias, ' !<>;:"[]+=-!@#$%^&*()?/.,{}\|', '')
		nSaveDataSessionID = THIS.DataSessionID
      
		IF VARTYPE(m.nDataSessionID) <> 'N' OR m.nDataSessionID < 1
			nDataSessionID = THIS.DataSessionID
		ENDIF
		
      IF nDataSessionID <> THIS.DataSessionID
			SET DATASESSION TO (m.nDataSessionID)
		ENDIF

		THIS.ClearErrors()                    && clear any existing errors
		THIS.ClearQueryOutput()

		oResultCollection = THIS.OnExecuteQuery(m.cSQL, m.cAlias)
      
		IF VARTYPE(m.oResultCollection) <> 'O' OR m.oResultCollection.Count == 0
			oResultCollection = .NULL.
		ENDIF
			
		IF m.nSaveDataSessionID <> THIS.DataSessionID
			SET DATASESSION TO (m.nSaveDataSessionID)
		ENDIF

		SELECT (m.nSelect)
		
		RETURN m.oResultCollection
	ENDFUNC
	
	FUNCTION ClearQueryOutput()
		THIS.QueryResultOutput = SPACE(0)
	ENDFUNC
   
	FUNCTION AddToQueryOutput(cMsg)
		THIS.QueryResultOutput = THIS.QueryResultOutput  + ;
		                         IIF(EMPTY(THIS.QueryResultOutput), SPACE(0), CHR(10) + CHR(10)) + m.cMsg
	ENDFUNC

	PROCEDURE CloseTable(cAlias)
		IF USED(m.cAlias)
			USE IN (m.cAlias)
		ENDIF
	ENDPROC

	** Abstract methods for populating collections
	FUNCTION OnGetAvailableServers(oServerCollection AS ServerCollection)
	ENDFUNC

	FUNCTION OnGetDatabases(oDatabaseCollection AS DatabaseCollection)
	ENDFUNC
	
	FUNCTION OnGetTables(oTableCollection AS TableCollection)
	ENDFUNC

	FUNCTION OnGetViews(oViewCollection AS ViewCollection)
	ENDFUNC

	FUNCTION OnGetStoredProcedures(oStoredProcCollection AS StoredProcCollection)
	ENDFUNC

	FUNCTION OnGetParameters(oParameterCollection AS ParameterCollection, cStoredProcName, cOwner)
	ENDFUNC

	FUNCTION OnGetFunctions(oFunctionCollection AS FunctionCollection)
	ENDFUNC

	FUNCTION OnGetFunctionParameters(oParameterCollection AS ParameterCollection, cFuncName, cOwner)
	ENDFUNC
	
	FUNCTION OnGetSchema(oColumnCollection AS ColumnCollection, cTableName, cOwner)
	ENDFUNC

	FUNCTION OnGetStoredProcedureDefinition(cStoredProcName, cOwner) AS String
		RETURN SPACE(0)
	ENDFUNC
	
	FUNCTION OnGetFunctionDefinition(cFunctionName, cOwner) AS String
		RETURN SPACE(0)
	ENDFUNC
	
	FUNCTION OnGetViewDefinition(cViewName, cOwner) AS String
		RETURN SPACE(0)
	ENDFUNC
	
	FUNCTION OnBrowseData(cTableName, cOwner)
	ENDFUNC

   * RAS 7-Jul-2011, separated out the call to the form from the setup of the data (OnBrowseData)
   FUNCTION OnBrowseForm(cTableName, cOwner)
   ENDFUNC

	FUNCTION OnNewTable()
	ENDFUNC

	FUNCTION OnDesignTable(cTableName, cOwner)
	ENDFUNC

	FUNCTION OnNewView()
	ENDFUNC

	FUNCTION OnDesignView(cViewName, cOwner)
	ENDFUNC

	FUNCTION OnRunStoredProcedure(cStoredProcName, cOwner, oParamList)
	ENDFUNC

	FUNCTION OnExecuteQuery(cSQL, cAlias)
	ENDFUNC

ENDDEFINE


* <summary>
*	Collection class which all other DataMgmt classes are
*	derived from.
* </summary>
DEFINE CLASS CCollection AS Collection
   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION AddEntity(tcName)
		RETURN .NULL.
	ENDFUNC

   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION Add(txItem, tcKey, teBefore, teAfter)
		* silently ignore any duplicate keys
		IF THIS.GetKey(m.tcKey) <> 0
			NODEFAULT
			RETURN .F.
		ENDIF
	ENDFUNC

   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION GetStruct(tcObjectType, tcName)
		LOCAL oStruct
		
		m.oStruct = CREATEOBJECT("Empty")
		AddProperty(m.oStruct, "Type", m.tcObjectType)
		AddProperty(m.oStruct, "Name", m.tcName)
		
		RETURN m.oStruct
	ENDFUNC
ENDDEFINE


* <summary>
*	Collection class for servers.
* </summary>
DEFINE CLASS ServerCollection AS CCollection
   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION AddEntity(tcName)
		LOCAL oStruct

		m.oStruct = THIS.GetStruct("Server", m.tcName)
		THIS.Add(m.oStruct, m.tcName)
		
		RETURN m.oStruct
	ENDFUNC
ENDDEFINE

* <summary>
*	Collection class for databases.
* </summary>
DEFINE CLASS DatabaseCollection AS CCollection
   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION AddEntity(tcName, tcPhysicalFile, tcConnectString, tcServer, tcDatabaseType, tcUser)
		LOCAL oStruct
		
		m.oStruct = THIS.GetStruct("Database", m.tcName)
		AddProperty(m.oStruct, "ConnectString", EVL(m.tcConnectString, SPACE(0)))
		AddProperty(m.oStruct, "PhysicalFile", EVL(m.tcPhysicalFile, SPACE(0)))
		AddProperty(m.oStruct, "Server", EVL(m.tcServer, SPACE(0)))
		AddProperty(m.oStruct, "State", 0)
		AddProperty(m.oStruct, "DatabaseType", EVL(m.tcDatabaseType, SPACE(0)))
		AddProperty(m.oStruct, "User", EVL(m.tcUser, SPACE(0)))
		AddProperty(m.oStruct, "Owner", SPACE(0))
		AddProperty(m.oStruct, "Size", 0)
		THIS.Add(m.oStruct, m.tcName)
		
		RETURN m.oStruct
	ENDFUNC
ENDDEFINE

* <summary>
*	Collection class for tables in a database.
* </summary>
DEFINE CLASS TableCollection AS CCollection
   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION AddEntity(tcName, tcPhysicalFile, tcOwner)
		LOCAL oStruct
		
		m.oStruct = THIS.GetStruct("Table", m.tcName)
		AddProperty(m.oStruct, "PhysicalFile", EVL(m.tcPhysicalFile, SPACE(0)))
		AddProperty(m.oStruct, "Owner", EVL(NVL(m.tcOwner, SPACE(0)), SPACE(0)))
		THIS.Add(m.oStruct, IIF(EMPTY(m.tcOwner), SPACE(0), m.tcOwner + '.') + m.tcName)
		
		RETURN m.oStruct
	ENDFUNC
ENDDEFINE

* <summary>
*	Collection class for views in a database.
* </summary>
DEFINE CLASS ViewCollection AS CCollection
   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION AddEntity(tcName, tcOwner)
		LOCAL oStruct
		
		m.oStruct = THIS.GetStruct("View", m.tcName)
		cOwner = EVL(NVL(m.tcOwner, SPACE(0)), SPACE(0))
		AddProperty(m.oStruct, "Owner", m.tcOwner)
		THIS.Add(m.oStruct, IIF(EMPTY(m.tcOwner), SPACE(0), m.tcOwner + '.') + m.tcName)
		
		RETURN m.oStruct
	ENDFUNC
ENDDEFINE

* <summary>
*	Collection class for stored procs in a database.
* </summary>
DEFINE CLASS StoredProcCollection AS CCollection
   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION AddEntity(tcName, tcOwner)
		LOCAL oStruct
		
		m.oStruct = THIS.GetStruct("StoredProc", m.tcName)

		cOwner = EVL(NVL(m.tcOwner, SPACE(0)), SPACE(0))
		AddProperty(m.oStruct, "Owner", m.tcOwner)
		THIS.Add(m.oStruct, IIF(EMPTY(m.tcOwner), SPACE(0), m.tcOwner + '.') + m.tcName)
		
		RETURN m.oStruct
	ENDFUNC
ENDDEFINE


* <summary>
*	Collection class for functions in a database.
* </summary>
DEFINE CLASS FunctionCollection AS CCollection
   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION AddEntity(tcName, tcOwner)
		LOCAL oStruct

		cOwner = EVL(NVL(m.tcOwner, SPACE(0)), SPACE(0))
		m.oStruct = THIS.GetStruct("Function", m.tcName)
		AddProperty(m.oStruct, "Owner", m.tcOwner)
		THIS.Add(m.oStruct, IIF(EMPTY(m.tcOwner), SPACE(0), m.tcOwner + '.') + m.tcName)

		RETURN m.oStruct
	ENDFUNC
ENDDEFINE


* <summary>
*	Collection class for columns in a table.
* </summary>
DEFINE CLASS ColumnCollection AS CCollection
   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION AddEntity(tcName, tcDataType, tnLength, tnDecimals, tlIsNullable, tcDefaultValue, tlPrimaryKey, tlIdentity)
		LOCAL oStruct

		m.oStruct = THIS.GetStruct("Column", m.tcName)
		AddProperty(m.oStruct, "DataType", m.tcDataType)
		AddProperty(m.oStruct, "Length", EVL(m.tnLength, 0))
		AddProperty(m.oStruct, "Decimals", EVL(m.tnDecimals, .NULL.))
		AddProperty(m.oStruct, "IsNullable", m.tlIsNullable)
		AddProperty(m.oStruct, "DefaultValue", NVL(EVL(m.tcDefaultValue, SPACE(0)), SPACE(0)))
		AddProperty(m.oStruct, "PrimaryKey", m.tlPrimaryKey)
		AddProperty(m.oStruct, "Identity", m.tlIdentity)
		THIS.Add(m.oStruct, m.tcName)

		RETURN m.oStruct
	ENDFUNC
ENDDEFINE


* <summary>
*	Collection class for parameters in a stored procedure.
* </summary>
DEFINE CLASS ParameterCollection AS CCollection
   * RAS 25-Jul-2008, renamed parameters to include the "t" scope indicator
   * Problem with entities that have a cName column, made rest to the standard
	FUNCTION AddEntity(tcName, tcDataType, tnLength, tnDecimals, tcDefaultValue, tnDirection, tcDefaultValue)
		LOCAL oStruct

		IF VARTYPE(m.tnDirection) <> 'N' OR EMPTY(m.tnDirection)
			m.tnDirection = PARAM_UNKNOWN
		ENDIF
		m.tcName = EVL(m.tcName, RETURN_VALUE_LOC)

		m.oStruct = THIS.GetStruct("Parameter", m.tcName)
		AddProperty(m.oStruct, "DataType", m.tcDataType)
		AddProperty(m.oStruct, "Length", EVL(NVL(m.tnLength, 0), 0))
		AddProperty(m.oStruct, "Decimals", EVL(m.tnDecimals, .NULL.))
		AddProperty(m.oStruct, "DefaultValue", EVL(m.tcDefaultValue, SPACE(0)))
		AddProperty(m.oStruct, "Direction", m.tnDirection)
		THIS.Add(m.oStruct, m.tcName)

		RETURN m.oStruct
	ENDFUNC
ENDDEFINE

*: EOF :*