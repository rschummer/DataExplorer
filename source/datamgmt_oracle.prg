* <summary>
*   Data Management for Oracle connections.
* </summary>
*
* RAS 3-Jul-2011, made pass for m.Variables
*

#include "DataExplorer.h"
#include "foxpro.h"
#include "adovfp.h"

#DEFINE ORACLE_SYS_OWNERS 'CTXSYS','MDSYS','OLAPSYS','ORDPLUGINS','ORDSYS','OUTLN',;
 'QS','QS_CBADM','QS_CS','QS_ES','QS_OS','QS_WS','SYS','SYSMAN','SYSTEM','WKSYS','WMSYS','XDB'


DEFINE CLASS OracleDatabaseMgmt AS ADODatabaseMgmt OF DataMgmt_ADO.prg
   lCriteriaSupported = .F.

   FUNCTION OnGetTables(oTableCollection AS TableCollection)
      LOCAL oRS AS ADODB.RecordSet
      LOCAL cOwner,cView,cSQLStr

   
      IF THIS.ADOOpen()
         cView = THIS.GetOracleView("TABLES")
         
         IF EMPTY(m.cView)
            RETURN
         ENDIF
         
         cSQLStr = THIS.GetOracleSQL(m.cView, "TABLE_NAME")
         oRS     = THIS.oADO.Execute(m.cSQLStr)
         
         DO WHILE NOT m.oRS.EOF()
            cOwner = IIF(THIS.nObjectLevel=1, THIS.UserName, m.oRS.Fields('OWNER').Value)
            THIS.AddEntity(m.oRS.Fields('TABLE_NAME').Value, SPACE(0), m.cOwner)
            m.oRS.MoveNext()
         ENDDO
         
         m.oRS.Close()
         THIS.ADOClose()

         THIS.LoadEntities(m.oTableCollection,3)
      ENDIF
   ENDFUNC

   FUNCTION OnGetViews(oViewCollection AS ViewCollection)
      LOCAL oRS AS ADODB.RecordSet
      LOCAL cOwner, cView, cSQLStr
            
      IF THIS.ADOOpen()
         cView = THIS.GetOracleView("VIEWS")
         
         IF EMPTY(m.cView)
            RETURN
         ENDIF
         
         cSQLStr = THIS.GetOracleSQL(cView,"VIEW_NAME,TEXT")
         oRS = THIS.oADO.Execute(m.cSQLStr)
         
         DO WHILE NOT m.oRS.EOF()
            cOwner = IIF(THIS.nObjectLevel=1, THIS.UserName, m.oRS.Fields('OWNER').Value)
            THIS.AddEntity(m.oRS.Fields('VIEW_NAME').Value, m.cOwner)
            m.oRS.MoveNext()
         ENDDO
         
         m.oRS.Close()
         THIS.ADOClose()

         THIS.LoadEntities(m.oViewCollection, 2)
      ENDIF
   ENDFUNC

   * RAS 25-Jul-2008, prefixed local memvars, renamed parameters to include "t" scope
   * Problem with entities that have a cName column, made rest to the standard
   FUNCTION OnGetStoredProcedures(toStoredProcCollection AS StoredProcCollection)
      LOCAL oRS AS ADODB.RecordSet
      LOCAL cOwner,cView,cSQLStrl,cName

      IF THIS.ADOOpen()
         m.cView=THIS.GetOracleView("OBJECTS")
         
         IF EMPTY(m.cView)
            RETURN
         ENDIF
         
         m.cSQLStr = THIS.GetOracleSQL(m.cView,"OBJECT_NAME,OBJECT_TYPE")
         m.cSQLStr = m.cSQLStr + IIF(THIS.nObjectLevel=2,[ AND ],[ WHERE ]) + [OBJECT_TYPE = 'PROCEDURE']            
         m.oRS = THIS.oADO.Execute(m.cSQLStr)
         
         DO WHILE !m.oRS.EOF()
            m.cName = oRS.Fields('OBJECT_NAME').Value
            m.cOwner = IIF(THIS.nObjectLevel=1, THIS.UserName, m.oRS.Fields('OWNER').Value)
            THIS.AddEntity(m.cName, m.cOwner)
            m.oRS.MoveNext()
         ENDDO
         
         m.oRS.Close()

         THIS.ADOClose()

         THIS.LoadEntities(m.toStoredProcCollection, 2)
      ENDIF
   ENDFUNC

   FUNCTION MapParameter(cSQLParamType, nPosition)
      LOCAL nParamType
      
      IF VARTYPE(m.cSQLParamType) <> 'C'
         nParamType = PARAM_UNKNOWN
      ELSE
         cSQLParamType = UPPER(ALLTRIM(m.cSQLParamType))
         
         DO CASE
            CASE m.nPosition = 0
               nParamType = PARAM_RETURNVALUE         
            CASE m.cSQLParamType == "IN"
               nParamType = PARAM_INPUT
            CASE m.cSQLParamType == "OUT"
               nParamType = PARAM_OUTPUT
            CASE m.cSQLParamType == "IN/OUT"
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
      LOCAL oRS AS ADODB.RecordSet
      LOCAL cView, cSQLStr 

      IF THIS.ADOOpen()
         cView=THIS.GetOracleView("ARGUMENTS")
         
         IF EMPTY(m.cView)
            RETURN
         ENDIF
         
         cSQLStr = [SELECT * FROM ] + m.cView + [ WHERE OBJECT_NAME = '] + m.cStoredProcName + [']
         
         IF THIS.nObjectLevel # 1
            cSQLStr = m.cSQLStr + [ AND OWNER = '] + m.cOwner + [']
         ENDIF
         
         oRS = THIS.oADO.Execute(m.cSQLStr)
         
         DO WHILE NOT m.oRS.EOF()
            oParameterCollection.AddEntity( ;
                 IIF(ISNULL(m.oRS.Fields("ARGUMENT_NAME").Value),"RETURN_VALUE", m.oRS.Fields("ARGUMENT_NAME").Value), ;
                 m.oRS.Fields("DATA_TYPE").Value, ;
                 m.oRS.Fields("DATA_LENGTH").Value, ;
                 m.oRS.Fields("DATA_PRECISION").Value, ;
                 m.oRS.Fields("DEFAULT_VALUE").Value, ;
                 THIS.MapParameter(m.oRS.Fields("IN_OUT").Value, m.oRS.Fields("POSITION").Value) ;
             )
            
            m.oRS.MoveNext()
         ENDDO
         
         m.oRS.Close()

         THIS.ADOClose()
      ENDIF
   ENDFUNC

   * Not supported by ADO, but is supported by our Oracle code
   * RAS 25-Jul-2008, prefixed local memvars, renamed parameters to include "t" scope
   * Problem with entities that have a cName column, made rest to the standard
   FUNCTION OnGetFunctions(toFunctionCollection AS FunctionCollection)
      LOCAL oRS AS ADODB.RecordSet
      LOCAL cView, cSQLStr, cOwner, cName

      IF THIS.ADOOpen()
         m.cView=THIS.GetOracleView("OBJECTS")
         
         IF EMPTY(m.cView)
            RETURN
         ENDIF
         
         m.cSQLStr = THIS.GetOracleSQL(m.cView,"OBJECT_NAME,OBJECT_TYPE")
         m.cSQLStr = m.cSQLStr + IIF(THIS.nObjectLevel=2,[ AND ],[ WHERE ]) + [OBJECT_TYPE = 'FUNCTION']
         m.oRS = THIS.oADO.Execute(m.cSQLStr)            
         
         DO WHILE !m.oRS.EOF()
            m.cName = m.oRS.Fields('OBJECT_NAME').Value
            m.cOwner = IIF(THIS.nObjectLevel=1, THIS.UserName, m.oRS.Fields('OWNER').Value)
            THIS.AddEntity(m.cName, m.cOwner, SPACE(0))
            m.oRS.MoveNext()                     
         ENDDO
         
         m.oRS.Close()
         
         THIS.ADOClose()
         THIS.LoadEntities(toFunctionCollection, 2)
      ENDIF
   ENDFUNC

   FUNCTION OnGetFunctionParameters(oParameterCollection AS ParameterCollection, cFunctionName AS String, cOwner AS String)
      LOCAL oRS AS ADODB.RecordSet
      LOCAL cView,cSQLStr
      
      IF THIS.ADOOpen()
         cView=THIS.GetOracleView("ARGUMENTS")
         
         IF EMPTY(m.cView)
            RETURN
         ENDIF
         
         cSQLStr = [SELECT * FROM ] + m.cView + [ WHERE OBJECT_NAME = '] + m.cFunctionName + [']
         
         IF THIS.nObjectLevel # 1
            cSQLStr = m.cSQLStr + [ AND OWNER = '] + m.cOwner + [']
         ENDIF
         
         oRS = THIS.oADO.Execute(m.cSQLStr)

         DO WHILE NOT m.oRS.EOF()
            oParameterCollection.AddEntity( ;
                 IIF(ISNULL(oRS.Fields("ARGUMENT_NAME").Value),"RETURN_VALUE", m.oRS.Fields("ARGUMENT_NAME").Value), ;
                 m.oRS.Fields("DATA_TYPE").Value, ;
                 m.oRS.Fields("DATA_LENGTH").Value, ;
                 m.oRS.Fields("DATA_PRECISION").Value, ;
                 m.oRS.Fields("DEFAULT_VALUE").Value, ;
                 THIS.MapParameter(m.oRS.Fields("IN_OUT").Value, m.oRS.Fields("POSITION").Value) ;
             )
            m.oRS.MoveNext()
         ENDDO
         
         m.oRS.Close()
         THIS.oADO.Close()
      ENDIF
   ENDFUNC

   FUNCTION OnGetSchema(oColumnCollection AS ColumnCollection, cTableName, cOwner)
      LOCAL oRS AS ADODB.RecordSet
      LOCAL cView,cSQLStr
      LOCAL cDType, nDLen, nDDec

      IF THIS.ADOOpen()
         cView=THIS.GetOracleView("TAB_COLUMNS")
         
         IF EMPTY(m.cView)
            RETURN
         ENDIF
         
         cSQLStr = [SELECT COLUMN_NAME, DATA_TYPE, DATA_PRECISION, DATA_LENGTH, DATA_SCALE, NULLABLE, DATA_DEFAULT FROM ] + ;
                   m.cView + [ WHERE TABLE_NAME = '] + m.cTableName + [']
         
         IF THIS.nObjectLevel#1
            cSQLStr = m.cSQLStr + [ AND OWNER = '] + m.cOwner + [']
         ENDIF
         
         oRS = THIS.oADO.Execute(m.cSQLStr)

         DO WHILE NOT m.oRS.EOF()
            cDType = m.oRS.Fields("DATA_TYPE").Value
            cDLen  = IIF(m.cDType == "NUMBER", m.oRS.Fields("DATA_PRECISION").Value, m.oRS.Fields("DATA_LENGTH").Value)
            cDDec  = IIF(m.cDType == "NUMBER", m.oRS.Fields("DATA_SCALE").Value, m.oRS.Fields("DATA_PRECISION").Value)
            
            oColumnCollection.AddEntity( ;
                 m.oRS.Fields("COLUMN_NAME").Value, ;  && column name
                 m.cDType, ;  && data type
                 NVL(m.cDLen,0), ;  && length
                 NVL(m.cDDec,0), ;  && numeric scale (decimals)
                 NVL(m.oRS.Fields("NULLABLE").Value, 'N') == 'Y', ;
                 NVL(m.oRS.Fields("DATA_DEFAULT").Value, SPACE(0)) ;
             )
            m.oRS.MoveNext()
         ENDDO
         
         oRS.Close()
         THIS.ADOClose()
      ENDIF
   ENDFUNC

   FUNCTION OnGetStoredProcedureDefinition(cStoredProcName, cOwner) AS String
      LOCAL cView
      LOCAL cSQLStr
      LOCAL oRS
      LOCAL cDefinition
      
      cDefinition = SPACE(0)
      IF THIS.ADOOpen()
         * Get Source -- only for users
         cView=THIS.GetOracleView("SOURCE")
         
         IF EMPTY(m.cView)
            RETURN SPACE(0)
         ENDIF
         
         cSQLStr = [SELECT * FROM ] + m.cView + ;
                   [ WHERE NAME = '] + m.cStoredProcName + [' AND TYPE = 'PROCEDURE'] + ;
                   IIF(THIS.nObjectLevel = 1, SPACE(0), [ AND OWNER = '] + m.cOwner + [']) +;
                   [ ORDER BY LINE]
         oRS = THIS.oADO.Execute(m.cSQLStr)
         
         DO WHILE NOT m.oRS.EOF()
            cDefinition = m.cDefinition + m.oRS.Fields('TEXT').Value
            m.oRS.MoveNext()
         ENDDO
         
         m.oRS.Close()   

         THIS.ADOClose()
      ENDIF

      RETURN cDefinition
   ENDFUNC
   
   FUNCTION OnGetFunctionDefinition(cFunctionName, cOwner) AS String
      LOCAL cDefinition
      LOCAL cView
      LOCAL oRS
      LOCAL cSQLStr
      
      cDefinition = SPACE(0)
      
      IF THIS.ADOOpen()
         * Get Source -- only for users
         cView = THIS.GetOracleView("SOURCE")
         
         IF EMPTY(m.cView)
            RETURN SPACE(0)
         ENDIF
         
         cSQLStr = [SELECT * FROM ] + m.cView + ;
                   [ WHERE NAME = '] + m.cFunctionName + [' AND TYPE = 'FUNCTION'] + ;
                   IIF(THIS.nObjectLevel = 1, SPACE(0), [ AND OWNER = '] + m.cOwner + [']) +;
                   [ ORDER BY LINE]
         oRS = THIS.oADO.Execute(cSQLStr)
         
         DO WHILE !oRS.EOF()
            cDefinition = m.cDefinition + m.oRS.Fields('TEXT').Value
            m.oRS.MoveNext()
         ENDDO
         
         m.oRS.Close()
      ENDIF
      
      THIS.ADOClose()

      RETURN m.cDefinition
   ENDFUNC

   FUNCTION OnGetViewDefinition(cViewName, cOwner) AS String
      LOCAL cDefinition,cView,oRS
      LOCAL cSQLStr
      
      cDefinition = SPACE(0)
      
      IF THIS.ADOOpen()
         cView=THIS.GetOracleView("VIEWS")
         
         IF EMPTY(m.cView)
            RETURN SPACE(0)
         ENDIF            
         
         cSQLStr = [SELECT VIEW_NAME,TEXT FROM ] + m.cView + ;
                   [ WHERE VIEW_NAME = '] + m.cViewName + ['] + ;
                   IIF(THIS.nObjectLevel = 1, SPACE(0), [ AND OWNER = '] + m.cOwner + ['])
         oRS = THIS.oADO.Execute(m.cSQLStr)            
         
         DO WHILE NOT m.oRS.EOF()
            cDefinition = m.oRS.Fields('TEXT').Value
            m.oRS.MoveNext()
         ENDDO
         
         m.oRS.Close()            
      ENDIF
      
      THIS.ADOClose()

      RETURN m.cDefinition
   ENDFUNC
   
   FUNCTION OnRunStoredProcedure(cStoredProcName, cOwner, oParamList)

      LOCAL cSQL
      LOCAL cValue
      LOCAL cParamList
      
      IF AT(';', m.cStoredProcName) > 0
         cStoredProcName = ALLTRIM(LEFT(m.cStoredProcName, AT(';', m.cStoredProcName) - 1))
      ENDIF

      cSQL = "CALL " + ["] + m.cStoredProcName + ["]
      IF VARTYPE(m.oParamList) == 'O'
         cParamList = SPACE(0)
         
         FOR i = 1 TO m.oParamList.Count
            IF INLIST(m.oParamList.Item(i).Direction, PARAM_INPUT, PARAM_INPUTOUTPUT, PARAM_OUTPUT)
               cValue = RTRIM(TRANSFORM(m.oParamList.Item(i).DefaultValue))
               IF ATC("text", m.oParamList.Item(i).DataType) > 0 OR ATC("char", m.oParamList.Item(i).DataType) > 0 
                  cValue     = IIF(UPPER(m.cValue)=="DEFAULT" OR UPPER(cValue)=="NULL", SPACE(0), m.cValue)
                  cParamList = m.cParamList + IIF(EMPTY(m.cParamList), SPACE(0), ',') + "'" + m.cValue + "'"
               ELSE
                  IF !EMPTY(m.cValue)
                     cValue     = IIF(UPPER(m.cValue)=="DEFAULT" OR UPPER(m.cValue)=="NULL", "null", m.cValue)
                     cParamList = m.cParamList + IIF(EMPTY(m.cParamList), SPACE(0), ',') + m.cValue
                  ENDIF
               ENDIF
            ENDIF
         ENDFOR
         IF !EMPTY(m.cParamList)
            cSQL = m.cSQL + '(' + m.cParamList + ')' 
         ENDIF
      ENDIF

      DO FORM RunQuery WITH THIS, m.cSQL, .T.
   ENDFUNC
   
   FUNCTION GetOracleView(cView)
      LOCAL oRS, lcView, lcViewExpr, lcSaveExact
      
      lcViewExpr  = SPACE(0)
      lcSaveExact = SET("EXACT")
      SET EXACT ON
      lcView = IIF(THIS.nObjectLevel=1,"USER","ALL") + "_" + m.cView
      
      TRY
         * Check to see if view exists...
         oRS = THIS.oADO.Execute("SELECT * FROM ALL_VIEWS WHERE VIEW_NAME = '&lcView'")
         * Check if any views exist that meet criteria
         DO WHILE NOT m.oRS.EOF()
            * It must be a system view
            IF INLIST(m.oRS.Fields('OWNER').Value, ORACLE_SYS_OWNERS)
               lcViewExpr = m.lcView
               EXIT
            ENDIF
            
            m.oRS.MoveNext()
         ENDDO
         
         m.oRS.Close()
      
      CATCH
      ENDTRY      
      
      SET EXACT &lcSaveExact
      RETURN m.lcViewExpr
   ENDFUNC

   FUNCTION GetOracleSQL(cView,cFields)
      LOCAL lcSQLStr
      lcSQLStr = "SELECT "
      
      DO CASE
         CASE m.cFields = "*"
            lcSQLStr = m.lcSQLStr + "*"
         
         CASE THIS.nObjectLevel = 1
            lcSQLStr = m.lcSQLStr + m.cFields
        
         OTHERWISE
            lcSQLStr = m.lcSQLStr + m.cFields + ",OWNER"
      ENDCASE
      
      lcSQLStr = m.lcSQLStr + " FROM " + m.cView
      
      IF THIS.nObjectLevel = 2
         lcSQLStr = m.lcSQLStr + " WHERE OWNER NOT IN (" + [ORACLE_SYS_OWNERS] + ")"
      ENDIF
      
      RETURN m.lcSQLStr
   ENDFUNC

   FUNCTION OnGetRunQuery(oCurrentNode)
      LOCAL cSQL, cStoredProcName, oParamList, cOwner, cParmStr, i

      IF TYPE("m.oCurrentNode.NodeData") == 'O'
         DO CASE
            CASE m.oCurrentNode.NodeData.TYPE == "StoredProc"
               cParmStr = SPACE(0)

               cStoredProcName = m.oCurrentNode.NodeData.Name
               cOwner          = m.oCurrentNode.NodeData.Owner
               oParamList      = THIS.GetParameters(m.cStoredProcName, m.cOwner)
               
               IF VARTYPE(m.oParamList) == 'O' AND m.oParamList.Count > 0
                  FOR i = 1 TO m.oParamList.Count
                     cParmStr = m.cParmStr + IIF(m.i=1, SPACE(0), ", ") + m.oParamList.Item[m.i].Name
                  ENDFOR
               ENDIF

               IF AT(SPACE(1), m.cStoredProcName) > 0
                  cStoredProcName = ["] + m.cStoredProcName +  ["]
               ENDIF

               RETURN "CALL " + m.cStoredProcName + "(" + m.cParmStr + ")"
         ENDCASE
      ENDIF
      
      RETURN SPACE(0)
   ENDFUNC

   FUNCTION SetObjectLevel(oNode)
      LOCAL nLevel
      
      IF VARTYPE(oNode)#"O"
         RETURN
      ENDIF
      
      DO FORM objectlevel WITH m.oNode
   ENDFUNC

ENDDEFINE

*: EOF :*