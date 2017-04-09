* <summary>
*	Data Management for Visual FoxPro.
* </summary>
*
* RAS 3-Jul-2011, made pass for m.Variables
*

#include "DataExplorer.h"
#include "foxpro.h"

DEFINE CLASS VFPDatabaseMgmt AS DatabaseMgmt OF DataMgmt.prg
	DatabaseFile = SPACE(0)
	
	FUNCTION DatabaseFile_ACCESS
		RETURN THIS.DatabaseFile
	ENDFUNC


	FUNCTION Connect(cDatabase, cUserName, cPassword) AS Boolean
		LOCAL lSuccess
		
		lSuccess = .F.

		IF VARTYPE(m.cUserName) <> 'C'
			cUserName = SPACE(0)
		ENDIF
		
      IF VARTYPE(m.cPassword) <> 'C'
			cPassword = SPACE(0)
		ENDIF
		
		
		* make sure we can get it open
		lSuccess = THIS.OpenDBC(m.cDatabase)
		
		* no need to leave it open once we've
		* determined we can open it
		IF lSuccess	
			THIS.CloseDBC()
		ENDIF
		
		RETURN m.lSuccess
	ENDFUNC
	
	FUNCTION OnGetTables(oTableCollection AS TableCollection)
		LOCAL i
		LOCAL nCnt
		LOCAL ARRAY aObjList[1]

		IF THIS.OpenDBC()
			nCnt = ADBOBJECTS(aObjList, "TABLE")

			IF m.nCnt > 0 AND THIS.SortObjects
				ASORT(aObjList)
			ENDIF
			
			FOR i = 1 TO m.nCnt
				oTableCollection.AddEntity(aObjList[i], ADDBS(JUSTPATH(DBC())) + DBGETPROP(aObjList[i], "TABLE", "Path"))
			ENDFOR
		
			THIS.CloseDBC()
		ENDIF
	ENDFUNC

	FUNCTION OnGetViews(oViewCollection AS ViewCollection)
		LOCAL i
		LOCAL nCnt
		LOCAL cViewDef
		LOCAL ARRAY aObjList[1]

		IF THIS.OpenDBC()
			nCnt = ADBOBJECTS(aObjList, "VIEW")

			IF m.nCnt > 0 AND THIS.SortObjects
				ASORT(aObjList)
			ENDIF
			
			FOR i = 1 TO m.nCnt
				* cViewDef = DBGETPROP(aObjList[i], "VIEW", "SQL")
				oViewCollection.AddEntity(aObjList[i], SPACE(0))
			ENDFOR

			THIS.CloseDBC()
		ENDIF
	ENDFUNC

	FUNCTION OnGetStoredProcedures(oStoredProcCollection AS StoredProcCollection)
		LOCAL i
		LOCAL nCnt
		LOCAL nSelect
		LOCAL oException
		LOCAL cTempFile
		LOCAL cSafety
		LOCAL nLineCnt
		LOCAL j
		LOCAL cProcDef
		LOCAL ARRAY aObjList[1]
		LOCAL ARRAY aProcCode[1]

		IF THIS.OpenDBC()
			nSelect = SELECT()
			
			SELECT 0
			
			TRY
				USE (DBC()) ALIAS DBCCursor SHARED AGAIN
				LOCATE FOR Objectname = "StoredProceduresSource"
				IF FOUND()
					cTempFile = ADDBS(SYS(2023)) + SYS(2015) + ".tmp"
					STRTOFILE(DBCCursor.code, m.cTempFile)
					
					nCnt = APROCINFO(aObjList, m.cTempFile)
					IF nCnt > 0 AND THIS.SortObjects
						ASORT(aObjList)
					ENDIF
					
					FOR i = 1 TO m.nCnt
						IF aObjList[i, 3] == 'Procedure'  && type
							oStoredProcCollection.AddEntity(aObjList[i, 1], SPACE(0))
						ENDIF
					ENDFOR
					
					cSafety = SET("SAFETY")
					SET SAFETY OFF
					ERASE (m.cTempFile)
					SET SAFETY &cSafety
				ENDIF

			CATCH TO m.oException
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			
         FINALLY
				IF USED("DBCCursor")
					USE IN DBCCursor
				ENDIF
				SELECT (m.nSelect)
			ENDTRY

			THIS.CloseDBC()
		ENDIF
	ENDFUNC

	FUNCTION OnGetFunctions(oFunctionCollection AS FunctionCollection)
	ENDFUNC
	
	FUNCTION OnGetStoredProcedureDefinition(cStoredProcName, cOwner) AS String
		LOCAL i
		LOCAL nCnt
		LOCAL nSelect
		LOCAL oException
		LOCAL cTempFile
		LOCAL cSafety
		LOCAL nLineCnt
		LOCAL j
		LOCAL cProcDef
		LOCAL ARRAY aObjList[1]
		LOCAL ARRAY aProcCode[1]

		cProcDef = SPACE(0)
		IF THIS.OpenDBC()
			nSelect = SELECT()
			
			SELECT 0
			
			TRY
				USE (DBC()) ALIAS DBCCursor SHARED AGAIN
				LOCATE FOR Objectname = "StoredProceduresSource"
				IF FOUND()
					cTempFile = ADDBS(SYS(2023)) + SYS(2015) + ".tmp"
					STRTOFILE(DBCCursor.code, m.cTempFile)
					
					nCnt = APROCINFO(aObjList, m.cTempFile)
					FOR i = 1 TO m.nCnt
						IF aObjList[i, 3] == 'Procedure' AND aObjList[i, 1] == m.cStoredProcName 
							* put code into array so we can extract the definition
							nLineCnt = ALINES(aProcCode, DBCCursor.code)

							* extract the definition
							cProcDef = SPACE(0)
							FOR j = aObjList[i, 2] TO IIF(m.i == nCnt, m.nLineCnt, aObjList[i + 1, 2] - 1)
								cProcDef = m.cProcDef + IIF(m.j == 1, SPACE(0), CHR(13) + CHR(10)) + aProcCode[j]
								IF INLIST(UPPER(ALLTRIM(aProcCode[m.j])), [** "END], "ENDPROC", "ENDFUNC")
									EXIT
								ENDIF
							ENDFOR
							EXIT
						ENDIF
					ENDFOR
					
					cSafety = SET("SAFETY")
					SET SAFETY OFF
					ERASE (m.cTempFile)
					SET SAFETY &cSafety
				ENDIF

			CATCH TO m.oException
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			FINALLY
				IF USED("DBCCursor")
					USE IN DBCCursor
				ENDIF
				SELECT (m.nSelect)
			ENDTRY

			THIS.CloseDBC()
		ENDIF

		RETURN m.cProcDef
	ENDFUNC
	
	FUNCTION OnGetFunctionDefinition(cFunctionName, cOwner) AS String
		* no function definitions in VFP
		RETURN SPACE(0)
	ENDFUNC

	FUNCTION OnGetViewDefinition(cViewName, cOwner) AS String
		LOCAL i
		LOCAL nCnt
		LOCAL cViewDef
		LOCAL ARRAY aObjList[1]

		cViewDef = SPACE(0)
		IF THIS.OpenDBC()
			nCnt = ADBOBJECTS(aObjList, "VIEW")

			FOR i = 1 TO nCnt
				IF aObjList[i] == m.cViewName
					cViewDef = DBGETPROP(aObjList[i], "VIEW", "SQL")
					EXIT
				ENDIF
			ENDFOR

			THIS.CloseDBC()
		ENDIF
		
		RETURN m.cViewDef
	ENDFUNC
	

	FUNCTION OnGetSchema(oColumnCollection AS ColumnCollection, cTableName, cOwner)
		LOCAL i
		LOCAL nCnt
		LOCAL nSelect
		LOCAL ARRAY aObjList[1]
      LOCAL lcField
      LOCAL lcDefaultValue

		nSelect = SELECT()

		IF THIS.OpenTable(m.cTableName, "TableCursor")
			nCnt = AFIELDS(aObjList, "TableCursor")

         *< RAS 29-Jul-2006, change reference from this.SortObjects to this.SortColumns
         *< IF nCnt > 0 AND THIS.SortObjects
			IF m.nCnt > 0 AND this.SortColumns
				ASORT(aObjList)
			ENDIF
         
			FOR i = 1 TO m.nCnt
            *< RAS 1-Jul-2006, corrected DefaultValue parameter (6th one). AFIELDS() does not capture the default
            *< value for the columns. Problem was caused by extraneous comma and a empty parameter.
				*< oColumnCollection.AddEntity(aObjList[i, 1], aObjList[i, 2], aObjList[i, 3], aObjList[i, 4], , aObjList[i, 5], aObjList[i, 9])
            
            lcField = JUSTSTEM(m.cTableName) + "." + ALLTRIM(aObjList[i, 1])

            * RAS 28-Jul-2006, the reason for the structured error handling has to do with free tables, they 
            * do not have a database container and do not have default values.
            TRY
               lcDefaultValue = NVL(DBGETPROP(m.lcField, "FIELD", "DefaultValue"), SPACE(0))
            
            CATCH TO m.loException
               lcDefaultValue = SPACE(0)
            
            FINALLY
            ENDTRY
            
            m.oColumnCollection.AddEntity(aObjList[i, 1], aObjList[i, 2], aObjList[i, 3], aObjList[i, 4], aObjList[i, 5], m.lcDefaultValue, aObjList[i, 9])
			ENDFOR
			
			IF USED("TableCursor")
				USE IN TableCursor
			ENDIF

			THIS.CloseDBC()
		ENDIF

		SELECT (m.nSelect)
			
	ENDFUNC

	
	FUNCTION OnBrowseData(cTableName, cOwner)
		LOCAL nSelect, ;
            llReturnVal
		
		nSelect = SELECT()
      lcAlias = JUSTSTEM(m.cTableName)

		IF THIS.OpenTable(m.cTableName, lcAlias)
			*< DO FORM BrowseForm
			*< THIS.CloseDBC()
         * RAS 3-Jul-2011, added buffering for VFP data.
         llReturnVal = .T.
		ENDIF

		SELECT (nSelect)
      
      RETURN m.llReturnVal
	ENDFUNC
   
   FUNCTION OnBrowseForm(cTableName, cOwner)
      LOCAL nSelect, ;
            loException, ;
            lError
      
      nSelect = SELECT()
         
      TRY 
         SELECT JUSTSTEM(m.cTableName)
         DO FORM BrowseForm2 WITH "vfp", this.lBrowseReadOnly
            
      CATCH TO loException
         this.SetError(m.loException.Message)
         lError = .T.

      FINALLY
         this.CloseDBC()

         SELECT (nSelect)

      ENDTRY

   ENDFUNC
   

	FUNCTION OnExecuteQuery(cSQL, cAlias)
		LOCAL oException
		LOCAL i
		LOCAL nErrorCnt
		LOCAL lError
		LOCAL cErrorMsg
		LOCAL cSQLNew
		LOCAL nCnt
		LOCAL ARRAY aSQLInfo[1]
		
		LOCAL ARRAY aErrorList[1]
		LOCAL ARRAY aAliasList[1]
      
      LOCAL lnOldSys3054, ;
            lcShowPlanSingle, ;
            lnShowPlanLevel, ;
            loDataExplorerEngine

      loDataExplorerEngine = NEWOBJECT("DataExplorerEngine", "DataExplorerEngine.prg")
      
      IF ISNULL(m.loDataExplorerEngine)
         lnShowPlanLevel = 11
      ELSE
         lnShowPlanLevel = INT(m.loDataExplorerEngine.ShowPlanParameter)
      ENDIF
      
      loDataExplorerEngine = .NULL.
      lcShowPlanSingle     = SPACE(0)
		nSelect              = SELECT()
		lError               = .F.
		oResultCollection    = .NULL.
		
		IF ATC("INTO CURSOR", m.cSQL) == 0 
			cSQL = m.cSQL + " INTO CURSOR ResultCursor"
		ENDIF
		
		cSQLNew = SPACE(0)
		nCnt    = ALINES(aSQLInfo, m.cSQL, .T.)
      
		FOR i = 1 TO m.nCnt
			cSQLNew = m.cSQLNew + SPACE(1) + aSQLInfo[i]
		ENDFOR

		IF EMPTY(THIS.DatabaseFile) OR THIS.OpenDBC()

			TRY
				* this will hold collection of aliases created by the query
				oResultCollection = CREATEOBJECT("Collection")
            
            * RAS 16-July-2006, added SQL Showplan details
            lnOldSys3054 = INT(VAL(SYS(3054)))
            
            SYS(3054, m.lnShowPlanLevel, "lcShowPlanSingle")

 				&cSQLNew

            SYS(3054, m.lnOldSys3054)

				oResultCollection.Add(ALIAS())

            * RAS 15-Jul-2006, added logic to display the number of records 
            * of the result set in query results.
            THIS.AddToQueryOutput(LOWER(ALIAS()) + ":" + SPACE(1) + TRANSFORM(RECCOUNT(ALIAS())) + ;
                                  SPACE(1) + ROWS_RETURNED ) 
            
            * RAS 16-Jul-2006, added logic to add ShowPlan details to query results
            IF EMPTY(m.lcShowPlanSingle)
               * Nothing to add to the query output
            ELSE
               THIS.AddToQueryOutput(m.lcShowPlanSingle)
            ENDIF

			CATCH TO oException
				THIS.SetError(m.oException.Message)
				lError = .T.

			FINALLY
			ENDTRY
			
			IF m.lError
				oResultCollection = .NULL.
				THIS.CloseDBC()
			ENDIF
         
			* THIS.CloseDBC()
		ENDIF
				
		SELECT (m.nSelect)
		
		
		RETURN m.oResultCollection
	ENDFUNC

	* Close all open tables in this session
	FUNCTION CloseAll()
		LOCAL i
		LOCAL nCnt
		LOCAL ARRAY aAliasList[1]
		
		nCnt = AUSED(aAliasList)
		FOR i = 1 TO m.nCnt
			USE IN (aAliasList[m.i, 2])
		ENDFOR
	ENDFUNC

	* Open the specified table in a new work area.
	* Upon leaving this method, the new work area
	* is selected
	FUNCTION OpenTable(cTableName, cAlias, lExclusive)
		LOCAL nSelect
		LOCAL oException
		LOCAL lSuccess
		LOCAL i
		LOCAL nCnt
		
		lSuccess = .F.
		nSelect  = SELECT()

		cAlias     = CHRTRAN(UPPER(EVL(m.cAlias, "TableCursor")), SPACE(1), '_')
		cTableName = UPPER(cTableName)
		

		* determine if table is already open
		nCnt = AUSED(aTablesInUse)
      
		FOR i = 1 TO m.nCnt
			IF aTablesInUse[i, 1] == m.cAlias AND ;
			 (UPPER(DBF(aTablesInUse[i, 1])) == m.cTableName OR ;
			  (CURSORGETPROP("Database", aTablesInUse[i, 1]) == THIS.DatabaseFile AND !EMPTY(THIS.DatabaseFile)))
				IF NOT m.lExclusive OR ISEXCLUSIVE(m.cAlias)
					SELECT (m.cAlias)
					lSuccess = .T.
				ENDIF
			ENDIF
		ENDFOR

		IF NOT m.lSuccess
			SELECT 0
			
			IF EMPTY(THIS.DatabaseFile)
				* determine if table is part of a DBC
				TRY
					IF m.lExclusive
						USE (m.cTableName) ALIAS (m.cAlias) EXCLUSIVE
					ELSE
						USE (m.cTableName) ALIAS (m.cAlias) SHARED
					ENDIF
					THIS.DatabaseFile = CURSORGETPROP("Database")
				
            CATCH TO m.oException
					* ignore error - it'll be caught below
					* MESSAGEBOX(oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC)
				ENDTRY
			ENDIF
			
			IF EMPTY(THIS.DatabaseFile) OR THIS.OpenDBC()
				* if table is part of a database, then
				* prepend the database name to the table name
				IF EMPTY(JUSTPATH(m.cTableName)) AND NOT EMPTY(THIS.DatabaseFile)
					cTableName = JUSTSTEM(THIS.DatabaseFile) + '!' + m.cTableName
				ENDIF
			
				TRY
					IF m.lExclusive
						USE (m.cTableName) ALIAS (m.cAlias) EXCLUSIVE
					ELSE
						USE (m.cTableName) ALIAS (m.cAlias) SHARED
					ENDIF
					lSuccess = .T.
				
            CATCH TO m.oException
					* MESSAGEBOX(oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC)
					* SELECT (nSelect)
				ENDTRY
			ENDIF
		ENDIF
		
		IF NOT m.lSuccess
			SELECT (m.nSelect)
		ENDIF
		
		RETURN m.lSuccess
	ENDFUNC
	
	* -- Methods specific to this Data Management provider
	PROTECTED FUNCTION OpenDBC(cDatabase)
		LOCAL lFirstOpen
		LOCAL lSuccess
		LOCAL oException
		
		lSuccess   = .F.
		lFirstOpen = NOT EMPTY(m.cDatabase)
		
      IF m.lFirstOpen
			IF EMPTY(JUSTEXT(m.cDatabase))
				cDatabase = FORCEEXT(m.cDatabase, "DBC")
			ENDIF
		ELSE
			cDatabase = THIS.DatabaseFile
		ENDIF

		
		TRY
			IF FILE(m.cDatabase)
				OPEN DATABASE (m.cDatabase) SHARED
				SET DATABASE TO (JUSTSTEM(m.cDatabase))
				
				IF m.lFirstOpen
					THIS.DatabaseFile = DBC()
				ENDIF
				
				lSuccess = .T.
			ENDIF
		
      CATCH TO m.oException
			MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			* ignore error -- we just know we couldn't get it open
		ENDTRY

		THIS.DatabaseName = THIS.DatabaseFile
	
		RETURN m.lSuccess
	ENDFUNC

	FUNCTION OnNewTable()
		LOCAL lSuccess

		lSuccess = .F.
		IF !EMPTY(THIS.DatabaseFile)
			IF THIS.OpenDBC()
				SET DATABASE TO (JUSTSTEM(THIS.DatabaseFile))
				lSuccess = .T.
			ENDIF
		ELSE
			lSuccess = .T.
		ENDIF
		IF m.lSuccess
			CREATE
		ENDIF

		THIS.CloseDBC()
	ENDFUNC

	FUNCTION OnDesignTable(cTableName, cOwner)
		LOCAL nSelect
		
		nSelect = SELECT()

		IF THIS.OpenTable(m.cTableName, JUSTSTEM(m.cTableName), .T.)
			MODIFY STRUCTURE
			THIS.CloseAll()
		ELSE
			IF THIS.OpenTable(m.cTableName, JUSTSTEM(m.cTableName), .F.)  && try again not exclusive
				MODIFY STRUCTURE
				THIS.CloseAll()
			ENDIF
		ENDIF
		SELECT (m.nSelect)
		
		THIS.CloseDBC()
	ENDFUNC

	FUNCTION OnNewView()
		LOCAL lSuccess

		lSuccess = .F.
		IF !EMPTY(THIS.DatabaseFile)
			IF THIS.OpenDBC()
				SET DATABASE TO (JUSTSTEM(THIS.DatabaseFile))
				lSuccess = .T.
			ENDIF
		ELSE
			lSuccess = .T.
		ENDIF
		
		IF m.lSuccess
			CREATE VIEW
		ENDIF

		THIS.CloseDBC()
	ENDFUNC

	FUNCTION OnDesignView(cViewName, cOwner)
		LOCAL nSelect
		
		nSelect = SELECT()

		IF THIS.OpenDBC()
			MODIFY VIEW (m.cViewName)
		ENDIF
		SELECT (m.nSelect)
		
		THIS.CloseDBC()
	ENDFUNC
	
	FUNCTION OnEditStoredProc(cProcName)
		* No way in VFP to edit a single stored proc
		LOCAL nSelect
		
		nSelect = SELECT()

		IF THIS.OpenDBC()
			MODIFY PROCEDURE
		ENDIF
		SELECT (m.nSelect)
		
		THIS.CloseDBC()
	ENDFUNC
	
	PROTECTED FUNCTION CloseDBC()
		THIS.CloseAll()

		CLOSE DATABASES
	ENDFUNC
	
	
ENDDEFINE

*: EOF :*