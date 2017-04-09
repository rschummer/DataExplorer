* <summary>
*	Data Management for ADO connections.
* </summary>
*
* RAS 3-Jul-2011, made pass for m.Variables
*

#include "DataExplorer.h"
#include "foxpro.h"
#include "adovfp.h"


DEFINE CLASS ADODatabaseMgmt AS DatabaseMgmt OF DataMgmt.prg
	PROTECTED lCriteriaSupported

	oADO = .NULL.  && reference to ADO Connection

	PromptPassword = .T.
	ConnectionString = SPACE(0)
	CustomPassword = .F.
	CustomConnection = .F.
	
	Owner = SPACE(0)
	UserID = SPACE(0)
	UserName = SPACE(0)

	ConnectTimeOut   = ADOCONNECT_TIMEOUT_DEFAULT 
	QueryTimeOut     = ADOQUERY_TIMEOUT_DEFAULT
	
	DBMSName = SPACE(0)
	DBMSVersion = SPACE(0)
	
	nObjectLevel = 1	&& 1-My, 2-User, 3-All (including System)

	lCriteriaSupported = .T.
	
	nEntityCnt = 0
	nEntityMax = 0
	DIMENSION aEntity[1, 2]
	
	PROCEDURE Init()
		THIS.oADO = CREATEOBJECT("ADODB.Connection")
	ENDPROC
	
	PROCEDURE Destroy()
		THIS.ADOClose()
		THIS.oADO = .NULL.
	ENDPROC

	PROCEDURE Disconnect()
		LOCAL oaADO AS ADODB.Connection

		IF VARTYPE(THIS.oADO) == 'O'
			THIS.ADOClose()
		ENDIF
		
      THIS.oADO = .NULL.
	ENDPROC	
	
	FUNCTION Connect(cConnectionString) AS Boolean
		LOCAL i
		LOCAL cPropName
		LOCAL lSuccess
		LOCAL oException
		
		THIS.ServerName   = SPACE(0)
		THIS.DatabaseName = SPACE(0)

		THIS.ConnectionString = m.cConnectionString

		TRY
			THIS.oADO.ConnectionString = m.cConnectionString
		
      CATCH TO m.oException
			MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			THIS.oADO.ConnectionString = SPACE(0)
		ENDTRY

		IF THIS.ADOOpen()
			FOR i = 0 TO (THIS.oADO.Properties.Count - 1)
				cPropName = UPPER(THIS.oADO.Properties(i).Name)

				DO CASE
   				CASE cPropName = "DATA SOURCE"
   					IF !ISNULL(THIS.oADO.Properties(i).Value)
   						THIS.ServerName = THIS.oADO.Properties(i).Value
   					ENDIF

   				CASE cPropName = "INITIAL CATALOG"
   					IF !ISNULL(THIS.oADO.Properties(i).Value)
   						THIS.DatabaseName = THIS.oADO.Properties(i).Value
   					ENDIF

   				CASE cPropName = "CURRENT CATALOG"
   					IF !ISNULL(THIS.oADO.Properties(i).Value)
   						THIS.DatabaseName = THIS.oADO.Properties(i).Value
   					ENDIF

   				CASE cPropName = "DBMS NAME"
   					IF !ISNULL(THIS.oADO.Properties(i).Value)
   						THIS.DBMSName = THIS.oADO.Properties(i).Value
   					ENDIF

   				CASE cPropName = "DBMS VERSION"
   					IF !ISNULL(THIS.oADO.Properties(i).Value)
   						THIS.DBMSVersion = THIS.oADO.Properties(i).Value
   					ENDIF

   				CASE cPropName = "USER NAME"
   					IF !ISNULL(EVL(THIS.oADO.Properties(i).Value, .NULL.))
   						THIS.UserName = THIS.oADO.Properties(i).Value
   					ENDIF

   				CASE cPropName = "USER ID"
   					IF !ISNULL(EVL(THIS.oADO.Properties(i).Value, .NULL.))
   						THIS.UserID = THIS.oADO.Properties(i).Value
   					ENDIF

   				CASE cPropName = "PROVIDER NAME"

				ENDCASE
			ENDFOR

			IF ATC("FOXPRO",THIS.DBMSName) # 0			
				THIS.lCriteriaSupported = .F.
			ENDIF

			THIS.ADOClose()
			lSuccess = .T.
		ENDIF
		RETURN lSuccess
	ENDFUNC

	FUNCTION SetConnValue(cConnectionString, cKeywords, cValue, lForce)
		LOCAL i
		LOCAL j
		LOCAL nKeyCnt
		LOCAL nCnt
		LOCAL cWord
		LOCAL cNewConnectionString
		LOCAL lFound
		LOCAL ARRAY aConnString[1]
		LOCAL ARRAY aKeyList[1]
		
		nKeyCnt = ALINES(aKeyList, m.cKeywords, 0, '|')
		
		lFound = .F.
		cNewConnectionString = SPACE(0)
		nCnt = ALINES(aConnString, m.cConnectionString, .T., ';')
		
      FOR i = 1 TO m.nCnt
			lFound = .F.
			IF m.nKeyCnt > 0 AND !EMPTY(aConnString[i])
				cWord = ALLTRIM(LOWER(CHRTRAN(aConnString[i], SPACE(1), SPACE(0))))
				
            IF ATC('=', m.cWord) > 0
					cWord = ALLTRIM(LEFT(m.cWord, ATC('=', m.cWord) - 1))
				ENDIF
				
				FOR j = 1 TO m.nKeyCnt
					IF m.cWord == LOWER(aKeyList[j])
						lFound = .T.
						EXIT
					ENDIF
				ENDFOR
				IF m.lFound
					cNewConnectionString = m.cNewConnectionString + IIF(EMPTY(m.cNewConnectionString), SPACE(0), ';') + ALLTRIM(LEFT(aConnString[i], ATC('=', aConnString[i]) - 1)) + '=' + m.cValue
					nKeyCnt = 0
				ENDIF
			ENDIF

			IF NOT m.lFound
				cNewConnectionString = m.cNewConnectionString + IIF(EMPTY(m.cNewConnectionString), SPACE(0), ';') + aConnString[i]
			ENDIF
		ENDFOR
		
		IF m.lForce AND !EMPTY(m.cValue) AND m.nKeyCnt > 0
			cNewConnectionString = m.cNewConnectionString + IIF(EMPTY(m.cNewConnectionString), SPACE(0), ';') + aKeyList[1] + '=' + m.cValue
		ENDIF
		
		RETURN m.cNewConnectionString
	ENDFUNC


	FUNCTION GetConnValue(cConnectionString, cKeywords)
		LOCAL i
		LOCAL j
		LOCAL nKeyCnt
		LOCAL nCnt
		LOCAL cWord
		LOCAL cValue
		LOCAL ARRAY aConnString[1]
		LOCAL ARRAY aKeyList[1]
		
		nKeyCnt = ALINES(aKeyList, m.cKeywords, 0, '|')
		cValue  = SPACE(0)
		nCnt    = ALINES(aConnString, cConnectionString, .T., ';')

		FOR i = 1 TO m.nCnt
			lFound = .F.
			
         IF m.nKeyCnt > 0 AND !EMPTY(aConnString[i])
				cWord = ALLTRIM(LOWER(CHRTRAN(aConnString[i], SPACE(1), SPACE(0))))
				IF ATC('=', m.cWord) > 0
					cWord = ALLTRIM(LEFT(m.cWord, ATC('=', m.cWord) - 1))
				ENDIF
				
				FOR j = 1 TO m.nKeyCnt
					IF cWord == LOWER(aKeyList[j])
						cValue = ALLTRIM(SUBSTR(aConnString[i], ATC('=', aConnString[i]) + 1))
						EXIT
					ENDIF
				ENDFOR
			ENDIF

			IF !EMPTY(m.cValue)
				EXIT
			ENDIF
		ENDFOR
		

		RETURN m.cValue
	ENDFUNC


	FUNCTION ADOOpen()
		LOCAL oException
		LOCAL lTryAgain
		LOCAL nUserIDTryAgain
		LOCAL cPassword
		LOCAL oRetValue
		LOCAL cUserID
		LOCAL cConnectionString
		LOCAL oDataLink
		LOCAL oADOConn
		LOCAL lShowDialog
		LOCAL lNoUserInConnString
		LOCAL ARRAY aConnString[1]

		cConnectionString = THIS.ConnectionString

		cPassword = THIS.UserPassword
		cUserID   = EVL(THIS.UserID, THIS.UserName)

		lTryAgain           = .T.
		nUserIDTryAgain     = 0
		lNoUserInConnString = ATC("user id=", m.cConnectionString) = 0 AND;
                            ATC("userid=", m.cConnectionString) = 0 AND;
                            ATC("uid=", m.cConnectionString) = 0

		DO WHILE m.lTryAgain
			lShowDialog = .F.

			IF THIS.CustomPassword
				cConnectionString = THIS.SetConnValue(m.cConnectionString, "password|pwd", m.cPassword, .T.)			
								
				* Need to handle scenario where User ID not specified. Each provider uses a different
				* naming convention for user id so let's try common ones without reprompting each time.
				* This is only an issue if user name is not included from original string.
				
				IF m.lNoUserInConnString
					nUserIDTryAgain = m.nUserIDTryAgain + 1
				ENDIF

				DO CASE
   				CASE nUserIDTryAgain = 1
   					cConnectionString = THIS.SetConnValue(m.cConnectionString, "userid", m.cUserID, .T.)
   				CASE nUserIDTryAgain = 2
   					cConnectionString = THIS.SetConnValue(m.cConnectionString, "user id", m.cUserID, .T.)
   				CASE nUserIDTryAgain = 3
   					cConnectionString = THIS.SetConnValue(m.cConnectionString, "uid", m.cUserID, .T.)
   				OTHERWISE
   					cConnectionString = THIS.SetConnValue(m.cConnectionString, "userid|uid", m.cUserID, .T.)
				ENDCASE
			ELSE
				cPassword = THIS.GetConnValue(m.cConnectionString, "password|pwd")
				cUserID   = THIS.GetConnValue(m.cConnectionString, "userid|uid")
			ENDIF

			TRY
				THIS.oADO.ConnectionString  = m.cConnectionString
				THIS.oADO.ConnectionTimeout = THIS.ConnectTimeout
				THIS.oADO.CommandTimeout    = THIS.QueryTimeout
				THIS.oADO.Open()
				lTryAgain = .F.
				
				IF THIS.CustomPassword
					THIS.UserPassword = m.cPassword
					THIS.UserID = m.cUserID
				ENDIF
				
			CATCH TO m.oException
				IF THIS.PromptPassword AND m.nUserIDTryAgain < 4 AND;
					(ATC("log", m.oException.Message) > 0 OR;
					 ATC("password", m.oException.Message) > 0 OR;
					 ATC("user", m.oException.Message) > 0)
					
					IF BETWEEN(m.nUserIDTryAgain,1,2)
						cConnectionString = THIS.ConnectionString
						THIS.CustomPassword = .T.
					ELSE
						DO FORM GetPassword WITH m.cUserID, IIF(ISNULL(m.cPassword),"",m.cPassword) TO m.oRetValue
                  
						IF VARTYPE(m.oRetValue) == 'O'
							cUserID = m.oRetValue.UserName
							cPassword = m.oRetValue.Password
							THIS.CustomPassword = .T.
							nUserIDTryAgain = 0
						ELSE
							lTryAgain = .F.
						ENDIF
					ENDIF	
				ELSE
					MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
					lShowDialog = .T.
				ENDIF
			ENDTRY

			IF m.lShowDialog
				TRY
					oDataLink = CREATEOBJECT("DataLinks")
					oADOConn  = CREATEOBJECT('ADODB.Connection')
					oADOConn.ConnectionString = m.cConnectionString
					oDataLink.PromptEdit(m.oADOConn)
					IF VARTYPE(m.oADOConn) == 'O' AND TYPE("m.oADOConn.ConnectionString") == 'C' AND !(m.oADOConn.ConnectionString == m.cConnectionString)
						THIS.CustomConnection = .T.
						cConnectionString = oADOConn.ConnectionString
						THIS.ConnectionString = m.cConnectionString
					ELSE
						lTryAgain = .F.
					ENDIF
				CATCH TO m.oException
					cConnectionString = SPACE(0)
				ENDTRY
			ENDIF
			
		ENDDO
		
		RETURN THIS.oADO.State == ADSTATEOPEN
	ENDFUNC
	
	FUNCTION ADOClose()
		IF TYPE("THIS.oADO") == 'O' AND !ISNULL(THIS.oADO) AND THIS.oADO.State == ADSTATEOPEN
			THIS.oADO.Close()
		ENDIF
	ENDFUNC
	
	
	* -- we do this intermediate array stuff so we can sort before
	* -- calling AddEntity method on the collections.
	PROCEDURE ClearEntities()
		THIS.nEntityCnt = 0
		THIS.nEntityMax = 0
		DIMENSION THIS.aEntity[1, 3]
	ENDPROC
   
	PROCEDURE AddEntity(cValue1, cValue2, cValue3)
		THIS.nEntityCnt = THIS.nEntityCnt + 1
		
      IF THIS.nEntityCnt > THIS.nEntityMax
			THIS.nEntityMax = THIS.nEntityMax + 100
			DIMENSION THIS.aEntity[THIS.nEntityMax, 3]
		ENDIF
		
      THIS.aEntity[THIS.nEntityCnt, 1] = m.cValue1
		THIS.aEntity[THIS.nEntityCnt, 2] = NVL(m.cValue2, SPACE(0))
		THIS.aEntity[THIS.nEntityCnt, 3] = NVL(m.cValue3, SPACE(0))
	ENDPROC

	PROCEDURE LoadEntities(oEntityCollection, nDimensions)
		LOCAL i
		
		IF THIS.nEntityCnt > 0
			DIMENSION THIS.aEntity[THIS.nEntityCnt, 3]
			=ASORT(THIS.aEntity, 1)
			
         IF VARTYPE(m.nDimensions) <> 'N' OR m.nDimensions == 0
				nDimensions = 1
			ENDIF
			
         DO CASE
   			CASE m.nDimensions == 1
   				FOR i = 1 TO THIS.nEntityCnt
   					oEntityCollection.AddEntity(THIS.aEntity[i, 1])
   				ENDFOR
               
   			CASE m.nDimensions == 2
   				FOR i = 1 TO THIS.nEntityCnt
   					oEntityCollection.AddEntity(THIS.aEntity[i, 1], THIS.aEntity[i, 2])
   				ENDFOR
               
   			CASE m.nDimensions == 3
   				FOR i = 1 TO THIS.nEntityCnt
   					oEntityCollection.AddEntity(THIS.aEntity[i, 1], THIS.aEntity[i, 2], THIS.aEntity[i, 3])
   				ENDFOR
               
			ENDCASE
		ENDIF
      
		THIS.ClearEntities()
	ENDPROC

	
	FUNCTION OpenSchema(cSchemaType, aCriteria)
		LOCAL oRS
		LOCAL lSuccess
		LOCAL oException
		
		lSuccess = .F.
		
      IF THIS.lCriteriaSupported AND PCOUNT() > 1
			TRY
				oRS = THIS.oADO.OpenSchema(m.cSchemaType, @aCriteria)
				lSuccess = .T.
			
         CATCH TO m.oException
				* THIS.lCriteriaSupported = .F. && don't try again if we got an error the first time
			ENDTRY
		ENDIF
		
		IF NOT m.lSuccess
			* try again without criteria because not all providers support it
			TRY
				oRS = THIS.oADO.OpenSchema(m.cSchemaType)
				lSuccess = .T.
			
         CATCH TO oException
				* ignore error because it's probably saying that provider does not support
				* the operation
				* MESSAGEBOX(oException.Message, MB_ICONEXCLAMATION)
			ENDTRY
		ENDIF

		RETURN m.oRS
	ENDFUNC

	
	FUNCTION OnGetTables(oTableCollection AS TableCollection)
		LOCAL oRS AS ADODB.RecordSet
		LOCAL cOwner
		LOCAL ARRAY aCriteria[4]
		
      aCriteria[1] = EVL(THIS.DatabaseName, .NULL.)
		aCriteria[2] = .NULL. && EVL(THIS.Owner, .NULL.)
		aCriteria[3] = .NULL.
		aCriteria[4] = "TABLE"
		
		IF THIS.ADOOpen()
			oRS = THIS.OpenSchema(ADSCHEMATABLES, @aCriteria)
			
         IF VARTYPE(m.oRS) == 'O'
				IF !THIS.lCriteriaSupported
					oRS.Filter = "TABLE_TYPE = 'TABLE'"
				ENDIF
				
            DO WHILE NOT m.oRS.EOF()
					cOwner = m.oRS.Fields('TABLE_SCHEMA').Value
					THIS.AddEntity(m.oRS.Fields('TABLE_NAME').Value, SPACE(0), m.cOwner)
					oRS.MoveNext()
				ENDDO
				
            m.oRS.Close()
			ENDIF
			
         THIS.ADOClose()
			THIS.LoadEntities(oTableCollection,3)
		ENDIF

	ENDFUNC

	FUNCTION OnGetViews(oViewCollection AS ViewCollection)
		LOCAL oRS AS ADODB.RecordSet
		LOCAL cOwner
		LOCAL ARRAY aCriteria[4]

		aCriteria[1] = EVL(THIS.DatabaseName, .NULL.)
		aCriteria[2] = EVL(THIS.Owner, .NULL.)
		aCriteria[3] = .NULL.
		aCriteria[4] = "VIEW"

		IF THIS.ADOOpen()
			oRS = THIS.OpenSchema(ADSCHEMATABLES, @aCriteria)
			IF VARTYPE(m.oRS) == 'O'
				IF NOT THIS.lCriteriaSupported
					oRS.Filter = "TABLE_TYPE = 'VIEW'"
				ENDIF
            
				DO WHILE NOT m.oRS.EOF()
					cOwner = m.oRS.Fields('TABLE_SCHEMA').Value
					THIS.AddEntity(m.oRS.Fields('TABLE_NAME').Value, m.cOwner)
					oRS.MoveNext()
				ENDDO
            
				m.oRS.Close()
			ENDIF
			THIS.ADOClose()
			THIS.LoadEntities(oViewCollection, 2)
		ENDIF

	ENDFUNC

	FUNCTION OnGetStoredProcedures(oStoredProcCollection AS StoredProcCollection)
		LOCAL oRS AS ADODB.RecordSet
		LOCAL cOwner
		LOCAL ARRAY aCriteria[4]

		aCriteria[1] = EVL(THIS.DatabaseName, .NULL.)
		aCriteria[2] = EVL(THIS.Owner, .NULL.)
		aCriteria[3] = .NULL.
		aCriteria[4] = .NULL.
		
      IF THIS.ADOOpen()
			oRS = THIS.OpenSchema(ADSCHEMAPROCEDURES, @aCriteria)
			IF VARTYPE(m.oRS) == 'O'
				DO WHILE NOT m.oRS.EOF()
					* Available fields:
					* 	PROCEDURE_CATALOG
					*	PROCEDURE_SCHEMA
					*	PROCEDURE_NAME
					*	PROCEDURE_TYPE
					*	PROCEDURE_DEFINITION
					*	DESCRIPTION
					*	DATE_CREATED
					
					* SQL Server returns Procedure names like this: 
					*	ProcNameA;0
					*	ProcNameB;1
					* 
					* so we need to strip off the group at the end
					*
					IF NOT LOWER(m.oRS.Fields('PROCEDURE_SCHEMA').Value) == "sys" AND;
					  ATC("Microsoft SQL Server", THIS.dbmsname) # 0
						THIS.AddEntity(m.oRS.Fields('PROCEDURE_NAME').Value, m.oRS.Fields('PROCEDURE_SCHEMA').Value)
					ENDIF
					
               m.oRS.MoveNext()
				ENDDO
            
				m.oRS.Close()
			ENDIF
         
			THIS.ADOClose()
   		THIS.LoadEntities(oStoredProcCollection, 2)
		ENDIF
	ENDFUNC
	
	* Given an ADO parameter type (in/out, etc) map it to our internal types
	FUNCTION MapParameter(nADOParamType)
		LOCAL nParamType
		
		IF VARTYPE(m.nADOParamType) <> 'N'
			nParamType = PARAM_UNKNOWN
		ELSE
			DO CASE
   			CASE nADOParamType == ADPARAMINPUT
   				nParamType = PARAM_INPUT
   			CASE nADOParamType == ADPARAMOUTPUT
   				nParamType = PARAM_OUTPUT
   			CASE nADOParamType == ADPARAMINPUTOUTPUT
   				nParamType = PARAM_INPUTOUTPUT
   			CASE nADOParamType == ADPARAMRETURNVALUE
   				nParamType = PARAM_RETURNVALUE
   			OTHERWISE
   				nParamType = PARAM_UNKNOWN
			ENDCASE
		ENDIF
				
		RETURN m.nParamType
	ENDFUNC

	FUNCTION OnGetParameters(oParameterCollection AS ParameterCollection, cStoredProcName AS String, cOwner AS String)
		LOCAL oRS AS ADODB.RecordSet
		LOCAL ARRAY aCriteria[4]
		LOCAL nParamType
		LOCAL lSuccess
		LOCAL nDirection
		
		aCriteria[1] = EVL(THIS.DatabaseName, .NULL.)
		aCriteria[2] = EVL(THIS.Owner, .NULL.)
		aCriteria[3] = m.cStoredProcName
		aCriteria[4] = .NULL.
		
		lSuccess = .F.
		IF THIS.ADOOpen()
			oRS = THIS.OpenSchema(ADSCHEMAPROCEDUREPARAMETERS, @aCriteria)
			
         IF VARTYPE(m.oRS) == 'O'
				lSuccess   = .T.
				oRS.Filter = "PROCEDURE_NAME = '" + m.cStoredProcName + "'"
				
            DO WHILE NOT m.oRS.EOF()
					nDirection = THIS.MapParameter(m.oRS.Fields("PARAMETER_TYPE").Value)
					
               IF nDirection <> PARAM_RETURNVALUE
						oParameterCollection.AddEntity( ;
   						  m.oRS.Fields("PARAMETER_NAME").Value, ;
   						  THIS.GetDataType(m.oRS.Fields("DATA_TYPE").Value), ;
   						  NVL(m.oRS.Fields("CHARACTER_MAXIMUM_LENGTH").Value, m.oRS.Fields("NUMERIC_PRECISION").Value), ;
   						  m.oRS.Fields("NUMERIC_SCALE").Value, ;
   						  SPACE(0), ;
   						  m.nDirection ;
						 )
					ENDIF

					m.oRS.MoveNext()
				ENDDO
				
            m.oRS.Close()
			ENDIF
			
         THIS.ADOClose()
		ENDIF
		
		RETURN m.lSuccess
	ENDFUNC

	* Not supported by ADO
	FUNCTION OnGetFunctions(oFunctionCollection AS FunctionCollection)
		RETURN .F.
	ENDFUNC
   
	FUNCTION OnGetFunctionParameters(oParameterCollection AS ParameterCollection, cFunctionName AS String, cOwner AS String)
		RETURN .F.
	ENDFUNC

	FUNCTION OnGetSchema(oColumnCollection AS ColumnCollection, cTableName, cOwner)
		LOCAL oRS AS ADODB.RecordSet
		LOCAL ARRAY aCriteria[3]
		aCriteria[1] = EVL(THIS.DatabaseName, .NULL.)
		aCriteria[2] = EVL(THIS.Owner, .NULL.)
		aCriteria[3] = m.cTableName

		IF THIS.ADOOpen()
			oRS = THIS.OpenSchema(ADSCHEMACOLUMNS, @aCriteria)
			
         IF VARTYPE(m.oRS) == 'O'
				IF NOT THIS.lCriteriaSupported
					oRS.Filter = "TABLE_NAME = '" + m.cTableName + "'"
				ENDIF
				
            DO WHILE NOT m.oRS.EOF()
					oColumnCollection.AddEntity( ;
   					  m.oRS.Fields("COLUMN_NAME").Value, ;  && column name
   					  THIS.GetDataType(m.oRS.Fields("DATA_TYPE").Value), ;  && data type
   					  NVL(m.oRS.Fields("NUMERIC_PRECISION").Value, ;
   					  NVL(m.oRS.Fields("CHARACTER_MAXIMUM_LENGTH").Value, m.oRS.Fields("DATETIME_PRECISION").Value)), ;  && length
   					  NVL(m.oRS.Fields("NUMERIC_SCALE").Value, 0), ;  && numeric scale
   					  .F., ;
   					  IIF(m.oRS.Fields("COLUMN_HASDEFAULT").Value, m.oRS.Fields("COLUMN_DEFAULT").Value, SPACE(0)) ;
					 )

					m.oRS.MoveNext()
				ENDDO
				
            m.oRS.Close()
			ENDIF
			
         THIS.ADOClose()
		ENDIF
	ENDFUNC

	FUNCTION OnGetStoredProcedureDefinition(cStoredProcName, cOwner) AS String
		LOCAL oRS
		LOCAL cDefinition
		LOCAL ARRAY aCriteria[4]

		aCriteria[1] = EVL(THIS.DatabaseName, .NULL.)
		aCriteria[2] = EVL(THIS.Owner, .NULL.)
		aCriteria[3] = .NULL.
		aCriteria[4] = .NULL.
		
		cDefinition = SPACE(0)
		
      IF THIS.ADOOpen()
			oRS = THIS.OpenSchema(ADSCHEMAPROCEDURES, @aCriteria)
			
         IF VARTYPE(m.oRS) == 'O'
				DO WHILE NOT m.oRS.EOF()
					IF m.oRS.Fields('PROCEDURE_NAME').Value == m.cStoredProcName AND NVL(m.oRS.Fields('PROCEDURE_SCHEMA').Value, SPACE(0)) == m.cOwner 
						cDefinition = NVL(m.oRS.Fields('PROCEDURE_DEFINITION').Value, SPACE(0))
						EXIT
					ENDIF
					
               m.oRS.MoveNext()
				ENDDO
				
            m.oRS.Close()
			ENDIF
			
         THIS.ADOClose()
		ENDIF

		RETURN m.cDefinition
	ENDFUNC
	
	FUNCTION OnGetFunctionDefinition(cFunctionName, cOwner) AS String
	ENDFUNC

	FUNCTION OnGetViewDefinition(cViewName, cOwner) AS String
	ENDFUNC

	PROTECTED FUNCTION GetDataType(nDataType)
		LOCAL cDataType
		
		DO CASE
		CASE m.nDataType = ADEMPTY
			cDataType = "Empty"
		CASE m.nDataType = ADTINYINT
			cDataType = "TinyInt"
		CASE m.nDataType = ADSMALLINT
			cDataType = "SmallInt"
		CASE m.nDataType = ADINTEGER
			cDataType = "Integer"
		CASE m.nDataType = ADBIGINT
			cDataType = "BigInt"
		CASE m.nDataType = ADUNSIGNEDTINYINT
			cDataType = "UnsignedTinyInt"
		CASE m.nDataType = ADUNSIGNEDSMALLINT
			cDataType = "UnsignedSmallInt"
		CASE m.nDataType = ADUNSIGNEDINT
			cDataType = "UnsignedInt"
		CASE m.nDataType = ADUNSIGNEDBIGINT
			cDataType = "UnsignedBigInt"
		CASE m.nDataType = ADSINGLE
			cDataType = "Single"
		CASE m.nDataType = ADDOUBLE
			cDataType = "Double"
		CASE m.nDataType = ADCURRENCY
			cDataType = "Currency"
		CASE m.nDataType = ADDECIMAL
			cDataType = "Decimal"
		CASE m.nDataType = ADNUMERIC
			cDataType = "Numeric"
		CASE m.nDataType = ADBOOLEAN
			cDataType = "Boolean"
		CASE m.nDataType = ADERROR
			cDataType = "Error"
		CASE m.nDataType = ADUSERDEFINED
			cDataType = "UserDefined"
		CASE m.nDataType = ADVARIANT
			cDataType = "Variant"
		CASE m.nDataType = ADIDISPATCH
			cDataType = "Dispatch"
		CASE m.nDataType = ADIUNKNOWN
			cDataType = "IUnknown"
		CASE m.nDataType = ADGUID
			cDataType = "GUID"
		CASE m.nDataType = ADDATE
			cDataType = "Date"
		CASE m.nDataType = ADDBDATE
			cDataType = "Date"
		CASE m.nDataType = ADDBTIME
			cDataType = "Time"
		CASE m.nDataType = ADDBTIMESTAMP
			cDataType = "TimeStamp"
		CASE m.nDataType = ADBSTR
			cDataType = "BStr"
		CASE m.nDataType = ADCHAR
			cDataType = "Char"
		CASE m.nDataType = ADVARCHAR
			cDataType = "VarChar"
		CASE m.nDataType = ADLONGVARCHAR
			cDataType = "LongVarChar"
		CASE m.nDataType = ADWCHAR
			cDataType = "WChar"
		CASE m.nDataType = ADVARWCHAR
			cDataType = "VarWChar"
		CASE m.nDataType = ADLONGVARWCHAR
			cDataType = "LongVarWChar"
		CASE m.nDataType = ADBINARY
			cDataType = "Binary"
		CASE m.nDataType = ADVARBINARY
			cDataType = "VarBinary"
		CASE m.nDataType = ADLONGVARBINARY
			cDataType = "LongVarBinary"
		CASE m.nDataType = ADCHAPTER
			cDataType = "Chapter"
		OTHERWISE
			cDataType = "Unknown"
		ENDCASE
		
		RETURN cDataType
	ENDFUNC
	
	
	FUNCTION OnBrowseData(cTableName, cOwner)
		LOCAL nSelect
		LOCAL nDataSessionID
		LOCAL oException
		LOCAL oRS AS ADODB.Recordset
		LOCAL oCA

		nSelect = SELECT()	
		cAlias = CHRTRAN(m.cTableName, ' !<>;:"[]+=-!@#$%^&*()?/.,{}\|', SPACE(0))
		
      IF ATC(SPACE(1), m.cTableName) > 0
			cTableName = ["] + m.cTableName + ["]
		ENDIF
		
		TRY
			IF THIS.ADOOpen()
				oRS = CREATEOBJECT("ADODB.RecordSet")
				oRS.Open(m.cTableName, THIS.oADO, ADOPENKEYSET, ADLOCKREADONLY, ADCMDTABLE)

				oCA = CREATEOBJECT("CursorAdapter")
				oCA.DataSourceType = "ADO"
				oCA.DataSource = m.oRS
				oCA.Alias = m.cAlias

				IF oCA.CursorFill(.F., .F., -1, m.oRS)
					SELECT (m.cAlias)
					DO FORM BrowseForm WITH .T.
					USE IN (m.cAlias)
				ENDIF
			ENDIF
		
      CATCH TO oException
			MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
		
      FINALLY
			THIS.ADOClose()
		ENDTRY

		SELECT (m.nSelect)
	ENDFUNC
	
	FUNCTION OnGetDatabases(oDatabaseCollection AS DatabaseCollection)
	ENDFUNC

	FUNCTION OnExecuteQuery(cSQL, cAlias)
		LOCAL cConnString
		LOCAL oException
		LOCAL nErrorCnt
		LOCAL lError
		LOCAL cErrorMsg
		LOCAL oRS
		LOCAL oCommand
		LOCAL oCA
		LOCAL ARRAY aErrorInfo[1]
		LOCAL ARRAY aErrorList[1]

		nSelect = SELECT()
		
		lError = .F.

		
		* this will hold collection of aliases created by the query
		oResultCollection = .NULL.

		TRY
			IF THIS.ADOOpen()
				oRS = CREATEOBJECT("ADODB.Command")
				* oRS.Open(cTableName, THIS.oADO, ADOPENKEYSET, ADLOCKREADONLY, ADCMDTABLE)
				
				oCommand = CREATEOBJECT("ADODB.Command")

				oCA = CREATEOBJECT("CursorAdapter")
				oCA.DataSourceType = "ADO"
				oCA.DataSource = CREATEOBJECT("ADODB.RecordSet")
				oCA.DataSource.CursorLocation = ADUSECLIENT
				oCA.DataSource.LockType = ADLOCKOPTIMISTIC
				oCA.DataSource.ActiveConnection = THIS.oADO
				
				oCA.Alias = "ResultCursor"
				oCA.SelectCmd = m.cSQL


				IF oCA.CursorFill(.F., .F., -1)
					oCA.CursorDetach()
					* this will hold collection of aliases created by the query
					oResultCollection = CREATEOBJECT("Collection")
					oResultCollection.Add(m.oCA.Alias)

					THIS.AddToQueryOutput(m.oCA.Alias + ": " + STRTRAN(RETRIEVE_COUNT_LOC, "##", TRANSFORM(RECCOUNT(m.oCA.Alias))))
				ELSE
					IF AERROR(aErrorInfo) > 0
						THROW aErrorInfo[2]
					ELSE
						THIS.AddToQueryOutput(QUERY_NORESULTS_LOC)
					ENDIF
				ENDIF
			ENDIF
		
		CATCH TO m.oException
			THIS.SetError(EVL(m.oException.UserValue, m.oException.Message))
			lError = .T.

		FINALLY
			THIS.ADOClose()
		ENDTRY
		
		IF lError
			THIS.AddToQueryOutput(THIS.LastError)
		ENDIF
		
		SELECT (m.nSelect)
		
		
		RETURN m.oResultCollection
	ENDFUNC
	
	PROCEDURE OnRunStoredProcedure(cStoredProcName, cOwner, oParamList)

		LOCAL cSQL
		LOCAL cValue
		LOCAL cParamList
		
		IF AT(';', m.cStoredProcName) > 0
			cStoredProcName = ALLTRIM(LEFT(m.cStoredProcName, AT(';', m.cStoredProcName) - 1))
		ENDIF

		cSQL = "EXECUTE " + m.cStoredProcName
		
		IF VARTYPE(m.oParamList) == 'O'
			cParamList = SPACE(0)
			FOR i = 1 TO m.oParamList.Count
				IF INLIST(m.oParamList.Item(i).Direction, PARAM_INPUT, PARAM_INPUTOUTPUT, PARAM_OUTPUT)
					cValue = RTRIM(TRANSFORM(m.oParamList.Item(i).DefaultValue))
					IF UPPER(m.cValue) == "DEFAULT" OR UPPER(m.cValue) == "NULL"
						cParamList = m.cParamList + IIF(EMPTY(m.cParamList), SPACE(1), ',') + oParamList.Item(i).Name + '=' + LOWER(m.cValue)
					ELSE
						IF INLIST(m.oParamList.Item(i).DataType, "nvarchar", "varchar", "nchar", "char", "text", "ntext")
							cParamList = m.cParamList + IIF(EMPTY(m.cParamList), ' ', ',') + m.oParamList.Item(i).Name + "='" + m.cValue + "'"
						ELSE
							IF !EMPTY(m.cValue)
								cParamList = m.cParamList + IIF(EMPTY(m.cParamList), ' ', ',') + m.oParamList.Item(i).Name + '=' + m.cValue
							ENDIF
						ENDIF
					ENDIF
				ENDIF
			ENDFOR
			IF !EMPTY(m.cParamList)
				cSQL = cSQL + SPACE(1) + m.cParamList
			ENDIF
		ENDIF

		DO FORM RunQuery WITH THIS, m.cSQL, .T.
	ENDPROC

	FUNCTION OnGetRunQuery(m.oCurrentNode)
		RETURN SPACE(0)
	ENDFUNC 
ENDDEFINE

*: EOF :*