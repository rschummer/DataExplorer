* <summary>
*	Data Management for Microsoft SQL Server.
* </summary>
*
* RAS 3-Jul-2011, made pass for m.Variables
*

#include "DataExplorer.h"
#include "foxpro.h"


* These defines are used by OnGetAvailableServers() method
#define SQL_HANDLE_ENV			1
#define SQL_HANDLE_DBC			2
#define SQL_ATTR_ODBC_VERSION	200
#define SQL_OV_ODBC3			3
#define SQL_SUCCESS				0
#define SQL_NEED_DATA			99
#define DEFAULT_RESULT_SIZE		2048
#define SQL_DRIVER_STR			"DRIVER=SQL SERVER";


DEFINE CLASS SQLDatabaseMgmt AS DatabaseMgmt OF DataMgmt.prg
	TrustedConnection = .T.

	SQLHandle = 0

	ConnectTimeOut   = CONNECT_TIMEOUT_DEFAULT  && 15 seconds
	QueryTimeOut     = QUERY_TIMEOUT_DEFAULT && wait 600 seconds by default for a query before timing out
	AutoTransactions = .T.
	DispWarnings     = .F.

   lAsync = .F.
   lBatch = .F.


	PROCEDURE Init()
		DODEFAULT()
	ENDPROC
	

	* Return an ODBC Connection string
	FUNCTION GetODBCConnectionString()
		LOCAL cConnString
 
      *< cConnString = [driver=] + SQL_ODBC_DRIVER + [;]
      cConnString = [driver=] + SQL_NATIVE_DRIVER + [;]
		IF !EMPTY(THIS.ServerName)
			cConnString = m.cConnString + [SERVER=] + THIS.ServerName + [;]
		ENDIF
		
      IF !EMPTY(THIS.DatabaseName)
			cConnString = m.cConnString + [DATABASE=] + THIS.DatabaseName + [;]
		ENDIF

		* if both are blank, this will be setup as a trusted connection
		IF THIS.TrustedConnection
			cConnString = m.cConnString + [Trusted_Connection=yes;]
		ELSE
			cConnString = m.cConnString + [UID=] + THIS.UserName + [;]
			cConnString = m.cConnString + [PWD=] + THIS.UserPassword + [;]
		ENDIF		

		cConnString = m.cConnString + [APP=] + APP_NAME + [;]

		RETURN m.cConnString
	ENDFUNC

	* Return a Connection string for SQL Namespace objects
	FUNCTION GetNamespaceConnectionString()
		LOCAL cConnString
 
		cConnString = SPACE(0)
		IF !EMPTY(THIS.ServerName)
			cConnString = m.cConnString + [SERVER=] + THIS.ServerName + [;]
		ENDIF
		IF !EMPTY(THIS.DatabaseName)
			cConnString = m.cConnString + [DATABASE=] + THIS.DatabaseName + [;]
		ENDIF

		* if both are blank, this will be setup as a trusted connection
		IF EMPTY(THIS.UserName) AND EMPTY(THIS.UserPassword)
			cConnString = m.cConnString + [Trusted_Connection=yes;]
		ELSE
			cConnString = m.cConnString + [UID=] + THIS.UserName + [;]
			cConnString = m.cConnString + [PWD=] + THIS.UserPassword + [;]
		ENDIF		

		RETURN m.cConnString
	ENDFUNC
	
	PROCEDURE Disconnect()
		IF THIS.SQLHandle > 0
			SQLDisconnect(THIS.SQLHandle)
		ENDIF
	ENDPROC	
	
	FUNCTION Connect(cServer, cDatabase, lTrustedConnection, cUserName, cPassword) AS Boolean
		LOCAL nSuccess
		LOCAL oLoginInfo
		LOCAL oException
		LOCAL cConnString
		LOCAL nDispLogin
		LOCAL lDispWarnings
		LOCAL nConnectTimeout

		IF VARTYPE(m.cServer) <> 'C'
			cServer = .NULL.
		ENDIF
		
		IF VARTYPE(m.cUserName) <> 'C'
			cUserName = THIS.UserName
		ENDIF
		IF VARTYPE(m.cPassword) <> 'C'
			cPassword = THIS.UserPassword
		ENDIF
		
		IF EMPTY(m.cServer)
			* no server specified, so nothing to connect to
			RETURN .T.
		ENDIF

		IF PCOUNT() < 3 OR VARTYPE(m.lTrustedConnection) <> 'L'
			lTrustedConnection = EMPTY(m.cUserName)
		ENDIF

		
		IF THIS.SQLHandle > 0 AND UPPER(THIS.ServerName) == UPPER(m.cServer) AND UPPER(THIS.DatabaseName) == UPPER(m.cDatabase) AND ;
		  THIS.TrustedConnection = m.lTrustedConnection AND UPPER(THIS.UserName) == UPPER(m.cUserName) AND ;
		  UPPER(THIS.UserPassword) == UPPER(m.cPassword)
			RETURN .T.
		ENDIF

		
		THIS.Disconnect()


		THIS.ServerName        = m.cServer
		THIS.DatabaseName      = m.cDatabase

		nSuccess = 0
      
		DO WHILE m.nSuccess = 0
			THIS.TrustedConnection = m.lTrustedConnection
			THIS.UserName          = m.cUserName
			THIS.UserPassword      = m.cPassword		

			nDispLogin      = SQLGETPROP(0, "DispLogin")
			lDispWarnings   = SQLGETPROP(0, "DispWarnings")
			nConnectTimeout = SQLGETPROP(0, "ConnectTimeout")

			cConnString = THIS.GetODBCConnectionString()
			this.ConnectionString = m.cConnString

			TRY
				SQLSETPROP(0, "DispLogin", DB_PROMPTNEVER)
				SQLSETPROP(0, "DispWarnings", .F.)
				SQLSETPROP(0, "ConnectTimeout", THIS.ConnectTimeout) 

				THIS.SQLHandle = SQLSTRINGCONNECT(m.cConnString, .T.)
			
         CATCH TO m.oException
				* ignore error
			
         FINALLY
				SQLSETPROP(0, "DispLogin", m.nDispLogin)
				SQLSETPROP(0, "DispWarnings", m.lDispWarnings)
				SQLSETPROP(0, "ConnectTimeout", m.nConnectTimeout)
			ENDTRY
         
			IF THIS.SQLHandle > 0
				SQLSETPROP(THIS.SQLHandle, "Asynchronous", .F.)
				SQLSETPROP(THIS.SQLHandle, "BatchMode", .T.)
				SQLSETPROP(THIS.SQLHandle, "IdleTimeout", 0) && never time out

				SQLSETPROP(THIS.SQLHandle, "QueryTimeout", THIS.QueryTimeout) 
				SQLSETPROP(THIS.SQLHandle, "Transactions", IIF(THIS.AutoTransactions, DB_TRANSAUTO, DB_TRANSMANUAL))
				SQLSETPROP(THIS.SQLHandle, "DispWarnings", THIS.DispWarnings) 

				nSuccess = 2
			ELSE
				IF m.nSuccess == 0
					DO FORM SQLConnectAs WITH m.lTrustedConnection, m.cUserName, m.cPassword TO m.oLoginInfo
					
               IF VARTYPE(oLoginInfo) == 'O'
						lTrustedConnection = m.oLoginInfo.TrustedConnection
						cUserName          = m.oLoginInfo.UserName
						cPassword          = m.oLoginInfo.Password
					ELSE
						nSuccess = 1
					ENDIF
				ENDIF
			ENDIF
		ENDDO
		
		RETURN (m.nSuccess == 2)
	ENDFUNC

	
	FUNCTION OnGetTables(oTableCollection AS TableCollection)
		LOCAL cSQL

		TEXT TO m.cSQL TEXTMERGE NOSHOW PRETEXT 7
			SELECT * FROM [<<THIS.DatabaseName>>].INFORMATION_SCHEMA.TABLES
			 WHERE TABLE_TYPE = 'BASE TABLE'
			 <<IIF(THIS.SortObjects, "ORDER BY [TABLE_NAME]", "")>>
		ENDTEXT

		IF NOT ISNULL(THIS.ExecuteQuery(m.cSQL, "SchemaCursor"))
			SELECT SchemaCursor
			SCAN ALL
				oTableCollection.AddEntity(RTRIM(SchemaCursor.Table_Name), SPACE(0), RTRIM(SchemaCursor.Table_Schema))  && Table_Schema = owner
			ENDSCAN			
		ENDIF
		THIS.CloseTable("SchemaCursor")
	ENDFUNC

	FUNCTION OnGetViews(oViewCollection AS ViewCollection)
		LOCAL cSQL
		
		TEXT TO m.cSQL TEXTMERGE NOSHOW PRETEXT 7
			SELECT * FROM [<<THIS.DatabaseName>>].INFORMATION_SCHEMA.VIEWS
			 <<IIF(THIS.SortObjects, "ORDER BY [TABLE_NAME]", "")>>
		ENDTEXT

		IF NOT ISNULL(THIS.ExecuteQuery(m.cSQL, "SchemaCursor"))
			SELECT SchemaCursor
			SCAN ALL
				oViewCollection.AddEntity(RTRIM(SchemaCursor.Table_Name), RTRIM(SchemaCursor.Table_Schema))
			ENDSCAN			
		ENDIF
		THIS.CloseTable("SchemaCursor")
	ENDFUNC

	FUNCTION OnGetStoredProcedures(oStoredProcCollection AS StoredProcCollection)
		LOCAL cSQL
		
      *< RAS 29-Sep-2006, per email from YAG, to fix error 
      *<  "Cannot sort a row of size 8295, which is greater than the allowable maximum of 8094."
      *<  David Fung implemented a change and blogged about: 
      *<  http://weblogs.foxite.com/davidfung/archive/2006/08/19/2275.aspx
      *<  because of a problem with Taiwan locale making INFORMATION_SCHEMA.ROUTINES records 
      *<  have a row size greater than 8094
      *< 
		*< TEXT TO cSQL TEXTMERGE NOSHOW PRETEXT 7
		*< 	SELECT * FROM [<<THIS.DatabaseName>>].INFORMATION_SCHEMA.ROUTINES
		*< 	 WHERE ROUTINE_TYPE = 'PROCEDURE'
		*< 	 <<IIF(THIS.SortObjects, "ORDER BY [ROUTINE_NAME]", "")>>
		*< ENDTEXT

      TEXT TO m.cSQL TEXTMERGE NOSHOW PRETEXT 7
         SELECT Routine_Name, Routine_Schema FROM [<<THIS.DatabaseName>>].INFORMATION_SCHEMA.ROUTINES
          WHERE ROUTINE_TYPE = 'PROCEDURE'
          <<IIF(THIS.SortObjects, "ORDER BY [ROUTINE_NAME]", "")>>
      ENDTEXT

		IF NOT ISNULL(THIS.ExecuteQuery(m.cSQL, "SchemaCursor"))
			SELECT SchemaCursor
			SCAN ALL
				oStoredProcCollection.AddEntity(RTRIM(SchemaCursor.Routine_Name), RTRIM(SchemaCursor.Routine_Schema))
			ENDSCAN
		ENDIF
		THIS.CloseTable("SchemaCursor")
	ENDFUNC

	FUNCTION MapParameter(cSQLParamType)
		LOCAL nParamType
		
		IF VARTYPE(m.cSQLParamType) <> 'C'
			nParamType = PARAM_UNKNOWN
		ELSE
			cSQLParamType = UPPER(ALLTRIM(m.cSQLParamType))
			
         DO CASE
   			CASE m.cSQLParamType == "IN"
   				nParamType = PARAM_INPUT
   			CASE m.cSQLParamType == "OUT"
   				nParamType = PARAM_OUTPUT
   			CASE m.cSQLParamType == "INOUT"
   				nParamType = PARAM_INPUTOUTPUT
   			CASE m.cSQLParamType == "RETURN"
   				nParamType = PARAM_RETURNVALUE
   			OTHERWISE
   				nParamType = PARAM_UNKNOWN
			ENDCASE
		ENDIF
				
		RETURN m.nParamType
	ENDFUNC

	FUNCTION OnGetParameters(oParameterCollection AS ParameterCollection, cStoredProcName AS String, cOwner AS String)
		LOCAL cSQL
		
		TEXT TO m.cSQL TEXTMERGE NOSHOW PRETEXT 7
			SELECT * FROM [<<THIS.DatabaseName>>].INFORMATION_SCHEMA.PARAMETERS
			 WHERE SPECIFIC_NAME = '<<m.cStoredProcName>>' AND SPECIFIC_SCHEMA = '<<m.cOwner>>'
			 ORDER BY [ORDINAL_POSITION]
		ENDTEXT

		IF NOT ISNULL(THIS.ExecuteQuery(m.cSQL, "SchemaCursor"))
			SELECT SchemaCursor
			SCAN ALL
				m.oParameterCollection.AddEntity( ;
				  RTRIM(SchemaCursor.Parameter_Name), ;
				  RTRIM(SchemaCursor.Data_Type), ;
				  NVL(SchemaCursor.Character_Maximum_Length, SchemaCursor.Numeric_Precision), ;
				  SchemaCursor.Numeric_Scale, ;
				  SPACE(0), ;
				  THIS.MapParameter(SchemaCursor.Parameter_Mode) ;
				 )
			ENDSCAN
		ENDIF
		THIS.CloseTable("SchemaCursor")
	ENDFUNC

	FUNCTION OnGetSchema(oColumnCollection AS ColumnCollection, cTableName, cOwner)
		LOCAL cSQL
		
      * RAS 29-Jul-2006, change reference from this.SortObjects to this.SortColumns
		TEXT TO m.cSQL TEXTMERGE NOSHOW PRETEXT 7
			SELECT * FROM [<<THIS.DatabaseName>>].INFORMATION_SCHEMA.COLUMNS 
			 WHERE TABLE_NAME = '<<m.cTableName>>' AND TABLE_SCHEMA = '<<m.cOwner>>'
			 <<IIF(this.SortColumns, "ORDER BY [COLUMN_NAME]", "ORDER BY [ORDINAL_POSITION]")>>
		ENDTEXT

		IF NOT ISNULL(THIS.ExecuteQuery(m.cSQL, "SchemaCursor"))
			SELECT SchemaCursor
			SCAN ALL
				m.oColumnCollection.AddEntity( ;
				  RTRIM(SchemaCursor.Column_Name), ;
				  RTRIM(SchemaCursor.Data_Type), ;
				  NVL(SchemaCursor.Character_Maximum_Length, SchemaCursor.Numeric_Precision), ;
				  SchemaCursor.Numeric_Scale, ;
				  LEFT(SchemaCursor.Is_Nullable, 1) == 'Y', ;
				  SchemaCursor.Column_Default ;
				 )
			ENDSCAN
		ENDIF
		THIS.CloseTable("SchemaCursor")
	ENDFUNC
	
	FUNCTION SPHelpText(cObjName) AS String
		LOCAL nSelect
		LOCAL lUnicode
		LOCAL cResult
		
		cResult = SPACE(0)
		
		nSelect = SELECT()
		
		TEXT TO m.cSQL TEXTMERGE NOSHOW PRETEXT 7
			sp_helptext @objname='<<cObjName>>'
		ENDTEXT

		IF NOT ISNULL(THIS.ExecuteQuery(m.cSQL, "HelpTextCursor"))
			SELECT HelpTextCursor
			SCAN ALL
				cResult = m.cResult + HelpTextCursor.Text
			ENDSCAN
			
         IF NOT (LEFT(STRCONV(m.cResult, 6),2) == "??")
				cResult = STRCONV(m.cResult, 6)
			ENDIF
		ENDIF
		
		THIS.CloseTable("HelpTextCursor")
		
		SELECT (m.nSelect)
		
		RETURN m.cResult
	ENDFUNC
	

	FUNCTION OnGetStoredProcedureDefinition(cStoredProcName, cOwner) AS String
		RETURN THIS.SPHelpText(m.cOwner + '.' + m.cStoredProcName)
	ENDFUNC
	
	FUNCTION OnGetFunctionDefinition(cFunctionName, cOwner) AS String
		RETURN THIS.SPHelpText(m.cOwner + '.' + m.cFunctionName)
	ENDFUNC

	FUNCTION OnGetViewDefinition(cViewName, cOwner) AS String
		RETURN THIS.SPHelpText(m.cOwner + '.' + m.cViewName)
	ENDFUNC


	FUNCTION OnGetFunctions(oFunctionCollection AS FunctionCollection)
		LOCAL cSQL
		
		TEXT TO m.cSQL TEXTMERGE NOSHOW PRETEXT 7
			SELECT * FROM [<<THIS.DatabaseName>>].INFORMATION_SCHEMA.ROUTINES
			 WHERE ROUTINE_TYPE = 'FUNCTION'
			 <<IIF(THIS.SortObjects, "ORDER BY [ROUTINE_NAME]", "")>>
		ENDTEXT

		IF NOT ISNULL(THIS.ExecuteQuery(m.cSQL, "SchemaCursor"))
			SELECT SchemaCursor
			SCAN ALL
				m.oFunctionCollection.AddEntity(RTRIM(SchemaCursor.Routine_Name), RTRIM(SchemaCursor.Routine_Schema))
			ENDSCAN			
		ENDIF
		THIS.CloseTable("SchemaCursor")
	ENDFUNC

	FUNCTION OnGetFunctionParameters(oParameterCollection AS ParameterCollection, cFuncName AS String, cOwner AS String)
		LOCAL cSQL
		
		TEXT TO m.cSQL TEXTMERGE NOSHOW PRETEXT 7
			SELECT * FROM [<<THIS.DatabaseName>>].INFORMATION_SCHEMA.PARAMETERS
			 WHERE 
			  SPECIFIC_NAME = '<<cFuncName>>' AND 
			  SPECIFIC_SCHEMA = '<<cOwner>>'
			 ORDER BY [ORDINAL_POSITION]
		ENDTEXT

		IF NOT ISNULL(THIS.ExecuteQuery(m.cSQL, "SchemaCursor"))
			SELECT SchemaCursor
			SCAN ALL FOR ALLTRIM(Parameter_Mode) = "IN"
				m.oParameterCollection.AddEntity( ;
				  RTRIM(SchemaCursor.Parameter_Name), ;
				  RTRIM(SchemaCursor.Data_Type), ;
				  NVL(SchemaCursor.Character_Maximum_Length, SchemaCursor.Numeric_Precision), ;
				  SchemaCursor.Numeric_Scale, ;
				  SPACE(0), ;
				  THIS.MapParameter(SchemaCursor.Parameter_Mode) ;
				 )
			ENDSCAN
         
			SCAN ALL FOR ALLTRIM(Parameter_Mode) = "OUT"
				m.oParameterCollection.AddEntity( ;
				  RTRIM(SchemaCursor.Parameter_Name), ;
				  RTRIM(SchemaCursor.Data_Type), ;
				  NVL(SchemaCursor.Character_Maximum_Length, SchemaCursor.Numeric_Precision), ;
				  SchemaCursor.Numeric_Scale, ;
				  SPACE(0), ;
				  THIS.MapParameter(SchemaCursor.Parameter_Mode) ;
				 )
			ENDSCAN
		ENDIF
		THIS.CloseTable("SchemaCursor")
	ENDFUNC

	
	FUNCTION OnBrowseData(cTableName, cOwner)
		LOCAL lnSelect
		LOCAL nDataSessionID
		LOCAL cConnString
		LOCAL oException
		LOCAL lAsync
		LOCAL lBatch
		LOCAL cOwner
      LOCAL lReadOnly
		
		lnSelect = SELECT()
		
		*< cAlias = CHRTRAN(m.cTableName, ' !<>;:"[]+=-!@#$%^&*()?/.,{}\|', SPACE(0))

		* retrieve data using SQL passthru
		* nDataSessionID = THIS.DataSessionID

		this.lAsync = SQLGETPROP(THIS.SQLHandle, "Asynchronous")
		this.lBatch = SQLGETPROP(THIS.SQLHandle, "BatchMode")

		SQLSETPROP(THIS.SQLHandle, "Asynchronous", .F.)
		SQLSETPROP(THIS.SQLHandle, 'BatchMode', .T.)

		*SET DATASESSION TO 1

*!*   		TRY
*!*   			*< RAS, 3-Jul-2011, new editable BrowseSQL form from Joel Leach handle query
*!*   			*< IF SQLEXEC(THIS.SQLHandle, "SELECT * FROM " + IIF(!EMPTY(cOwner), "[" + cOwner + "].", '') + "[" + cTableName + "]", cAlias) >= 0
*!*   			*<    DO FORM BrowseForm WITH .T.
*!*      		*< ENDIF
*!*            DO FORM BrowseForm2 WITH "sql", this.lBrowseReadOnly, this.SQLHandle, m.cOwner, m.cTableName
*!*   		
*!*         CATCH TO m.oException
*!*   			MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
*!*   		
*!*         FINALLY
*!*   			* SET DATASESSION TO (nDataSessionID)

*!*   			SQLSETPROP(THIS.SQLHandle, "asynchronous", m.lAsync)
*!*   			SQLSETPROP(THIS.SQLHandle, 'BatchMode', m.lBatch)
*!*   		
*!*   			IF USED(m.cAlias)
*!*   				USE IN (m.cAlias)
*!*   			ENDIF
*!*   		ENDTRY

		SELECT (m.lnSelect)
	ENDFUNC

   FUNCTION OnBrowseForm(cTableName, cOwner)
      LOCAL lnSelect, ;
            loException, ;
            lError
      
      lnSelect = SELECT()
      lcAlias  = CHRTRAN(m.cTableName, ' !<>;:"[]+=-!@#$%^&*()?/.,{}\|', SPACE(0))
         
      TRY 
         DO FORM BrowseForm2 WITH "sql", this.lBrowseReadOnly, this.SQLHandle, m.cOwner, m.cTableName
         *DO FORM BrowseFormSQL WITH this.SQLHandle, m.cOwner, m.cTableName 
           
      CATCH TO loException
         this.SetError(m.loException.Message)
         lError = .T.

      FINALLY
         * SET DATASESSION TO (nDataSessionID)

         SQLSETPROP(THIS.SQLHandle, "Asynchronous", this.lAsync)
         SQLSETPROP(THIS.SQLHandle, 'BatchMode', this.lBatch)
      
         IF USED(m.lcAlias)
            USE IN (m.lcAlias)
         ENDIF

         SELECT (lnSelect)

      ENDTRY

   ENDFUNC
   
   
	FUNCTION OnExecuteQuery(cSQL, cAlias)
		LOCAL cConnString
		LOCAL oException
		LOCAL i
		LOCAL j
		LOCAL nErrorCnt
		LOCAL nAliasCnt
		LOCAL nSetNum
		LOCAL nResultCnt
		LOCAL nMoreResults
		LOCAL lError
		LOCAL cErrorMsg
		LOCAL lAsync
		LOCAL lBatch
		LOCAL ARRAY aErrorList[1]
		LOCAL ARRAY aAliasList[1]
		LOCAL ARRAY aCountInfo[1]

		nSelect = SELECT()
		
		lError = .F.

		
		* this will hold collection of aliases created by the query
		oResultCollection = CREATEOBJECT("Collection")

		nSetNum = 0
		nAliasCnt = AUSED(aAliasList)
      
		FOR i = 1 TO m.nAliasCnt
			FOR j = 1 TO LEN(aAliasList[m.i, 1])
				* find first digit and strip off the number
				* for example, "Sqlresult5" -> break into "Sqlresult" and 5
				IF ISDIGIT(SUBSTR(aAliasList[m.i, 1], m.j, 1))
					IF UPPER(LEFT(aAliasList[m.i, 1], m.j - 1)) == UPPER(m.cAlias)
						nSetNum = MAX(SUBSTR(aAliasList[m.i, 1], m.j), m.nSetNum)
					ENDIF
					EXIT
				ENDIF
			ENDFOR
		ENDFOR
		

		* retrieve data using SQL passthru
		lAsync = SQLGETPROP(THIS.SQLHandle, "asynchronous")
		lBatch = SQLGETPROP(THIS.SQLHandle, "batchmode")
		
		SQLSETPROP(THIS.SQLHandle, "asynchronous", .F.)
		SQLSETPROP(THIS.SQLHandle, "BatchMode", .F.)
		
      TRY
			nResultCnt = SQLEXEC(THIS.SQLHandle, m.cSQL, m.cAlias + IIF(m.nSetNum > 0, TRANSFORM(m.nSetNum), SPACE(0)), aCountInfo)
			THIS.ParseQueryResults(@aCountInfo)
			IF m.nResultCnt > 0
				DO WHILE .T.
					IF NOT EMPTY(aCountInfo[1, 1])
						m.oResultCollection.Add(m.cAlias + IIF(m.nSetNum > 0, TRANSFORM(m.nSetNum), SPACE(0)))
					ENDIF

					nSetNum = m.nSetNum + 1
					nMoreResults = SQLMORERESULTS(THIS.SQLHandle, m.cAlias + TRANSFORM(m.nSetNum), aCountInfo) 

					IF m.nMoreResults == 2
						EXIT
					ENDIF

					THIS.ParseQueryResults(@aCountInfo)
				ENDDO
			ENDIF

			IF m.nResultCnt < 0
				THIS.SetError()
				lError = .T.
			ENDIF
         
		CATCH TO m.oException
			THIS.SetError(m.oException.Message)
			lError = .T.

		FINALLY
			SQLSETPROP(THIS.SQLHandle, "asynchronous", m.lAsync)
			SQLSETPROP(THIS.SQLHandle, 'BatchMode', m.lBatch)
			
		ENDTRY
		
		IF m.lError
			* oResultCollection = .NULL.
			THIS.AddToQueryOutput(THIS.LastError)
		ENDIF
		
		SELECT (m.nSelect)
		
		
		RETURN m.oResultCollection
	ENDFUNC
	
	* Show results to output
	FUNCTION ParseQueryResults(aCountInfo)
		LOCAL cMsg
		LOCAL i

		FOR i = 1 TO ALEN(aCountInfo, 1)
			DO CASE
   			CASE EMPTY(aCountInfo[i, 1])
   				* no result set (INSERT, UPDATE, or DELETE)
   				THIS.AddToQueryOutput(QUERY_NORESULTS_LOC)
   			
            CASE aCountInfo[i, 1] == '0'
   				* no records or command failed
   			
            OTHERWISE
   				THIS.AddToQueryOutput(aCountInfo[i, 1] + ": " + IIF(aCountInfo[i, 2] = -1, "error", STRTRAN(RETRIEVE_COUNT_LOC, "##", TRANSFORM(aCountInfo[i, 2]))))
			ENDCASE
		ENDFOR
	ENDFUNC

	* Populate collection with available SQL servers
	FUNCTION OnGetAvailableServers(oServerCollection AS ServerCollection)
		LOCAL hEnv
		LOCAL hConn
		LOCAL cInString
		LOCAL cOutString
		LOCAL nLenOutString
		LOCAL ARRAY aServerList[1]

		DECLARE SHORT SQLBrowseConnect IN odbc32 ; 
		    INTEGER   ConnectionHandle, ; 
		    STRING    InConnectionString, ; 
		    INTEGER   StringLength1, ; 
		    STRING  @ OutConnectionString, ; 
		    INTEGER   BufferLength, ; 
		    INTEGER @ StringLength2Ptr
		    
		DECLARE SHORT SQLAllocHandle IN odbc32 ; 
		    INTEGER   HandleType, ; 
		    INTEGER   InputHandle, ; 
		    INTEGER @ OutputHandlePtr 
		    
		DECLARE SHORT SQLFreeHandle IN odbc32 ; 
		    INTEGER HandleType, ; 
		    INTEGER Handle 

		DECLARE SHORT SQLSetEnvAttr IN odbc32 ; 
		    INTEGER EnvironmentHandle, ; 
		    INTEGER Attribute, ; 
		    INTEGER ValuePtr, ; 
		    INTEGER StringLength 


		hEnv = 0
		hConn = 0
		cInString = SQL_DRIVER_STR
		cOutString = SPACE(DEFAULT_RESULT_SIZE)
		nLenOutString = 0

		TRY
			IF SQLAllocHandle(SQL_HANDLE_ENV, m.hEnv, @hEnv) == SQL_SUCCESS
				IF (SQLSetEnvAttr(m.hEnv, SQL_ATTR_ODBC_VERSION, SQL_OV_ODBC3, 0)) == SQL_SUCCESS
					IF SQLAllocHandle(SQL_HANDLE_DBC, m.hEnv, @hConn) == SQL_SUCCESS
						IF (SQLBrowseConnect(m.hConn, @cInString, LEN(m.cInString), @cOutString, DEFAULT_RESULT_SIZE, @nLenOutString)) == SQL_NEED_DATA
							nCnt = ALINES(aServerList, STREXTRACT(m.cOutString, '{', '}'), .T., ',')
							FOR i = 1 TO m.nCnt
								oServerCollection.AddEntity(aServerList[m.i])
							ENDFOR
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		
      CATCH TO oException
			* ignore error, just return an empty collection of servers
		
      FINALLY
			IF m.hConn <> 0
				SQLFreeHandle(SQL_HANDLE_DBC, m.hConn)
			ENDIF
			
         IF m.hEnv <> 0
				SQLFreeHandle(SQL_HANDLE_ENV, m.hConn)
			ENDIF
		ENDTRY
	ENDFUNC

	FUNCTION OnGetDatabases(oDatabaseCollection AS DatabaseCollection)
		LOCAL oException
		LOCAL oDatabase

		TRY
			IF !ISNULL(THIS.ExecuteQuery("sp_helpdb", "SchemaCursor"))
				SELECT SchemaCursor
				SCAN ALL
					oDatabase = oDatabaseCollection.AddEntity(RTRIM(SchemaCursor.Name))
					oDatabase.Owner = RTRIM(SchemaCursor.Owner)
					oDatabase.Size = SchemaCursor.db_size
				ENDSCAN
			ENDIF
		
      CATCH TO m.oException
			MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
		ENDTRY
		
      THIS.CloseTable("SchemaCursor")
	ENDFUNC


	FUNCTION OnRunStoredProcedure(cStoredProcName, cOwner, oParamList)
		LOCAL cSQL
		LOCAL cValue
		LOCAL cParamList
		
		IF LEFT(m.cStoredProcName, 1) <> '['
			cStoredProcName = '[' + m.cStoredProcName + ']'
		ENDIF
		
      cSQL       = "EXECUTE " + IIF(EMPTY(m.cOwner), SPACE(0), "[" + m.cOwner + "].") + m.cStoredProcName
		cParamList = SPACE(0)

		FOR i = 1 TO m.oParamList.Count
			IF INLIST(m.oParamList.Item(i).Direction, PARAM_INPUT, PARAM_INPUTOUTPUT, PARAM_OUTPUT)
				cValue = RTRIM(TRANSFORM(m.oParamList.Item(i).DefaultValue))
				IF UPPER(m.cValue) == "DEFAULT" OR UPPER(m.cValue) == "NULL"
					cParamList = m.cParamList + IIF(EMPTY(m.cParamList), SPACE(1), ',') + m.oParamList.Item(i).Name + '=' + LOWER(m.cValue)
				ELSE
					IF INLIST(m.oParamList.Item(i).DataType, "nvarchar", "varchar", "nchar", "char", "text", "ntext")
						cParamList = m.cParamList + IIF(EMPTY(m.cParamList), SPACE(1), ',') + m.oParamList.Item(i).Name + "='" + m.cValue + "'"
					ELSE
						IF !EMPTY(cValue)
							cParamList = m.cParamList + IIF(EMPTY(m.cParamList), SPACE(1), ',') + m.oParamList.Item(i).Name + '=' + m.cValue
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDFOR
		IF !EMPTY(m.cParamList)
			cSQL = m.cSQL + SPACE(1) + m.cParamList
		ENDIF

		DO FORM RunQuery WITH THIS, m.cSQL, .T.
	ENDFUNC


	FUNCTION GenerateScripts(cTableName)
		IF EMPTY(EVL(m.cTableName, SPACE(0)))
			THIS.ShowSQLDialog("Generate Scripts", 'Database')
		ELSE
			THIS.ShowSQLDialog("Generate Scripts", 'Table', m.cTableName)
		ENDIF
	ENDFUNC


	* Display SQL Enterprise Manager dialog
	* 	<cCmdName> = namespace command to execute (e.g. "Generate Scripts")
	* 	[cObjType] = {"Server", "Database", "Table", "View", "Procedure", "Function"}
	*
	* Examples:
	*	o.ShowSQLDialog("Generate Scripts", 'Table', "EISFAC")
	*	o.ShowSQLDialog("Generate Scripts", 'D')
	*
	FUNCTION ShowSQLDialog(xCmdName, cObjType, cObjName)
		LOCAL oSQLNS
		LOCAL oServer
		LOCAL oDatabases
		LOCAL oDatabase
		LOCAL oTables
		LOCAL oTable
		LOCAL oViews
		LOCAL oView
		LOCAL oStoredProcs
		LOCAL oStoreProc
		LOCAL oNSObject
		LOCAL oRootObject

		cObjType = UPPER(EVL(m.cObjType, "Server"))
		
		oSQLNS = CREATEOBJECT("SQLNS.SQLNamespace")
		oSQLNS.Initialize(DATAEXPLORER_LOC, SQLNSROOTTYPE_SERVER, THIS.GetNameSpaceConnectionString())

		* drill in until we get to the object we want to operate on		
		oServer = oSQLNS.GetRootItem()

		IF cObjType == 'SERVER'  && server
			oRootObject = m.oServer
		ELSE
			oDatabases = oSQLNS.GetFirstChildItem(m.oServer, SQLNSOBJECTTYPE_DATABASES)
			oDatabase  = oSQLNS.GetFirstChildItem(m.oDatabases, SQLNSOBJECTTYPE_DATABASE, THIS.DatabaseName)
			
			IF m.cObjType == 'DATABASE' && database
				oRootObject = m.oDatabase
			ELSE
				DO CASE
				CASE m.cObjType == 'TABLES'
					oRootObject = m.oSQLNS.GetFirstChildItem(m.oDatabase, SQLNSOBJECTTYPE_DATABASE_TABLES)

				CASE m.cObjType == 'VIEWS' && view
					oRootObject = m.oSQLNS.GetFirstChildItem(m.oDatabase, SQLNSOBJECTTYPE_DATABASE_VIEWS)

				CASE m.cObjType == 'PROCEDURES' && stored procedure
					oRootObject = m.oSQLNS.GetFirstChildItem(m.oDatabase, SQLNSOBJECTTYPE_DATABASE_SPS)

				CASE m.cObjType == 'FUNCTIONS' && function
					oRootObject = m.oSQLNS.GetFirstChildItem(m.oDatabase, SQLNSOBJECTTYPE_DATABASE_UDFS)

				CASE m.cObjType == 'TABLE' && table
					oTables = m.oSQLNS.GetFirstChildItem(m.oDatabase, SQLNSOBJECTTYPE_DATABASE_TABLES)
					oTable  = m.oSQLNS.GetFirstChildItem(m.oTables, SQLNSOBJECTTYPE_DATABASE_TABLE, cObjName)

					oRootObject = m.oTable

				CASE m.cObjType == 'VIEW' && view
					oViews = m.oSQLNS.GetFirstChildItem(m.oDatabase, SQLNSOBJECTTYPE_DATABASE_VIEWS)
					oView  = m.oSQLNS.GetFirstChildItem(m.oViews, SQLNSOBJECTTYPE_DATABASE_VIEW, cObjName)

					oRootObject = m.oView

				CASE m.cObjType == 'PROCEDURE' && stored procedure
					oStoredProcs = m.oSQLNS.GetFirstChildItem(m.oDatabase, SQLNSOBJECTTYPE_DATABASE_SPS)
					oStoredProc  = m.oSQLNS.GetFirstChildItem(m.oStoredProcs, SQLNSOBJECTTYPE_DATABASE_SP, cObjName)

					oRootObject = m.oStoredProc

				CASE m.cObjType == 'FUNCTION' && function
					oFunctions = m.oSQLNS.GetFirstChildItem(m.oDatabase, SQLNSOBJECTTYPE_DATABASE_UDFS)
					oFunction  = m.oSQLNS.GetFirstChildItem(m.oStoredProcs, SQLNSOBJECTTYPE_DATABASE_UDF, cObjName)

					oRootObject = m.oStoredProc
				ENDCASE
			ENDIF
		ENDIF
		
		oNSObject = m.oSQLNS.GetSQLNamespaceObject(m.oRootObject)
		
		IF VARTYPE(m.xCmdName) == 'N'
			m.oNSObject.ExecuteCommandByID(m.xCmdName, _SCREEN.Hwnd)
		ELSE	
			m.oNSObject.ExecuteCommandByName(m.xCmdName, _SCREEN.Hwnd)
		ENDIF
	ENDFUNC	

ENDDEFINE

*: EOF :*