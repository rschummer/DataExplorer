* <summary>
* 	Default node for populating tree.  All view nodes
*	should be derived from INode.
*
*	For each created node, override the OnPopulate() method
*	to load child nodes.
* </summary>
*
*  RAS 3-Jul-2011, made pass for m.Variables
*

#include "foxpro.h"
#include "DataExplorer.h"


DEFINE CLASS INode AS Custom
	PROTECTED lRetrieved
	
	oNodeList = .NULL.

	lRetrieved = .F.  && set to true when sub nodes have been retrieved

	NodeText  = ''
	NodeID    = ''
	NodeType  = ''  && contains UniqueID of connection type
	NodeData  = .NULL.
	NodeKey   = ''
	ImageKey  = ''  && uniqueID of image to display next to node
	ExpandOnInit = .F.

	* populated automatically
	NodeLevel  = 0
	NodeOrder  = 0
	ParentID   = ''
	ParentNode = .NULL.
	SaveNode   = .F.  && set to TRUE if this node can be persisted
	
	EndNode   = .F.
	Expanded  = .F.

	Inactive  = .F.
	
	Options = ''
	OptionData = ''
	oOptionCollection  = .NULL.
	
	DefType = ''
	
	
	* set this in DetailTemplate_Access
	DetailTemplate = '' && ex: "<row><caption>#NodeText#</caption><value></value></row>"

	Filtered = .F.	


	* Return .T. if given method name exists on node
	* Used by menus
	FUNCTION IsOkay(cMethodName)
		LOCAL lOkay
		
		lOkay = .F.
		IF PEMSTATUS(THIS, m.cMethodName, 5)
			IF PEMSTATUS(THIS, m.cMethodName + "Okay", 5)
				lOkay = EVALUATE("THIS." + m.cMethodName + "Okay()")
			ENDIF
		ENDIF
		
		RETURN m.lOkay
	ENDFUNC

	* retrieve a DataExplorerEngine configuration value
	FUNCTION GetConfigValue(cProperty, xDefault)
		LOCAL xValue
		LOCAL oDataExplorerEngine

		xValue = .NULL.
		
		oDataExplorerEngine = NEWOBJECT("DataExplorerEngine", "DataExplorerEngine.prg")
		oDataExplorerEngine.RestorePrefs()
		
      IF PEMSTATUS(m.oDataExplorerEngine, m.cProperty, 5)
			xValue = EVALUATE("oDataExplorerEngine." + m.cProperty)
		ELSE
			IF PCOUNT() > 1
				xValue = m.xDefault
			ENDIF
		ENDIF
		
      oDataExplorerEngine = .NULL.

		RETURN m.xValue
	ENDFUNC


	PROCEDURE Init(cNodeText, xNodeData, cNodeKey)
		IF EMPTY(THIS.NodeID)
			THIS.NodeID = "user." + SYS(2015)
		ENDIF

		IF VARTYPE(m.cNodeText) == 'C'
			THIS.NodeText = m.cNodeText
		ENDIF
		
      IF PCOUNT() > 1
			THIS.NodeData = m.xNodeData
		ENDIF

		THIS.NodeKey = EVL(NVL(m.cNodeKey, SPACE(0)), SPACE(0))

		THIS.oNodeList         = CREATEOBJECT("Collection")
		THIS.oOptionCollection = CREATEOBJECT("Collection")

		* regular expression filter that gets applied to all objects
		THIS.CreateOption("FilterInclude", SPACE(0))
		THIS.CreateOption("FilterExclude", SPACE(0))

		THIS.OnInit()
	ENDFUNC

	FUNCTION Filtered_ACCESS()
		RETURN !EMPTY(THIS.GetOption("FilterInclude", SPACE(0))) OR NOT EMPTY(THIS.GetOption("FilterExclude", SPACE(0)))
	ENDFUNC

	PROCEDURE OnInit()
	ENDPROC
	
	PROCEDURE Destroy()
		THIS.ReleaseNodes()
	ENDPROC

	* -- Override either one of these to customize
	* -- the node text to return
	FUNCTION NodeText_ACCESS
		RETURN THIS.OnGetNodeText()
	ENDFUNC
	FUNCTION OnGetNodeText()
		RETURN THIS.NodeText
	ENDFUNC

	FUNCTION NodeKey_ACCESS
		RETURN EVL(THIS.NodeKey, THIS.OnGetNodeText())
	ENDFUNC


	FUNCTION DetailTemplate_ACCESS
		RETURN SPACE(0) && THIS.DetailTemplate
	ENDFUNC
	
	FUNCTION ImageKey_ACCESS
		RETURN THIS.ImageKey
	ENDFUNC

	* Populate node with children
	FUNCTION OnPopulate()
		RETURN .T.
	ENDFUNC

	PROTECTED PROCEDURE CopyIntoNodeData(oDestination, oSource)
		LOCAL i
		LOCAL nCnt
		LOCAL aPropList[1]
		
		oDestination.NodeData = CREATEOBJECT("Empty")
		
		IF VARTYPE(m.oSource) == 'O'
			nCnt = AMEMBERS(aPropList, m.oSource, 0, 'G')
			FOR i = 1 TO m.nCnt
				ADDPROPERTY(m.oDestination.NodeData, aPropList[i], EVALUATE("m.oSource." + aPropList[i]))
			ENDFOR
		ENDIF
	ENDPROC
	
	* create an instance of a node, and assign any options
	* that have the same name in the parent to have the
	* same value.
	FUNCTION CreateNode(cClassName, cClassLib, cNodeText, xNodeData, cNodeKey)
		LOCAL oNode
		LOCAL i
		LOCAL cFullClassLib

		IF EMPTY(m.cClassLib)
			oNode = CREATEOBJECT(m.cClassName)
		ELSE
			* if classlib is not in the path, then look for it
			cFullClassLib = m.cClassLib
			
         IF EMPTY(JUSTEXT(m.cFullClassLib))
				cFullClassLib = FORCEEXT(m.cFullClassLib, "prg")
			ENDIF
			
         IF !FILE(m.cFullClassLib)
				cFullClassLib = FORCEPATH(m.cFullClassLib, HOME(7))
				IF !FILE(m.cFullClassLib)
					cFullClassLib = FORCEPATH(m.cFullClassLib, HOME())
				ENDIF
			ENDIF
			
         oNode = NEWOBJECT(m.cClassName, m.cFullClassLib)
		ENDIF
		
      IF PCOUNT() > 2 AND VARTYPE(m.cNodeText) == 'C'
			oNode.NodeText = m.cNodeText
		ENDIF
		
      IF PCOUNT() > 3
			* copy all properties from xNodeData.
			* don't assign the reference directly because causes
			* problem with leaving the data session open that
			* the original node data was created in
			THIS.CopyIntoNodeData(m.oNode, m.xNodeData)
		ENDIF
		
      oNode.NodeKey = EVL(NVL(m.cNodeKey, SPACE(0)), SPACE(0))

		IF TYPE("THIS.DataMgmtClass") == 'C' AND TYPE("m.oNode.DataMgmtClass") == 'C'
			oNode.DataMgmtClass        = THIS.DataMgmtClass
			oNode.DataMgmtClassLibrary = THIS.DataMgmtClassLibrary
		ENDIF
		
      IF TYPE("THIS.ProviderName") == 'C' AND TYPE("m.oNode.ProviderName") == 'C'
			oNode.ProviderName = THIS.ProviderName
		ENDIF
		
		FOR i = 1 TO THIS.oOptionCollection.Count
			IF !INLIST(THIS.oOptionCollection.Item(i).OptionName, "FilterInclude", "FilterExclude")
				oNode.SetOption(THIS.oOptionCollection.Item(i).OptionName, THIS.oOptionCollection.Item(i).OptionValue)
			ENDIF
		ENDFOR

		RETURN m.oNode
	ENDFUNC



	* Called to retrieve information on node to
	* display in Details area
	FUNCTION OnGetDetails() AS String
		LOCAL i
		LOCAL nCnt
		LOCAL cDetails
		LOCAL cValue
		LOCAL ARRAY aNodeInfo[1]

		cDetails = THIS.DetailTemplate
		
      IF !EMPTY(m.cDetails)
			IF VARTYPE(THIS.NodeData) == 'O'
				nCnt = AMEMBERS(aNodeInfo, THIS.NodeData, 0, 'U')
				
            FOR i = 1 TO nCnt
					cValue   = TRANSFORM(NVL(EVALUATE("THIS.NodeData." + aNodeInfo[i]), SPACE(0)))
					cDetails = STRTRAN(m.cDetails, "#Node." + aNodeInfo[i] + '#', m.cValue, -1, -1, 1)
					cDetails = STRTRAN(m.cDetails, "#" + aNodeInfo[i] + '#', m.cValue, -1, -1, 1)
				ENDFOR
			ENDIF

			nCnt = AMEMBERS(aNodeInfo, THIS, 0, 'U')
			
         FOR i = 1 TO nCnt
				cValue   = TRANSFORM(NVL(EVALUATE("THIS." + aNodeInfo[i]), SPACE(0)))
				cDetails = STRTRAN(cDetails, "#Node." + aNodeInfo[i] + '#', m.cValue, -1, -1, 1)
				cDetails = STRTRAN(cDetails, "#" + aNodeInfo[i] + '#', m.cValue, -1, -1, 1)
			ENDFOR

			* cDetails = TEXTMERGE(cDetails, .F., "<<", ">>")
		ENDIF
		
		RETURN m.cDetails
		
	ENDFUNC

	FUNCTION GetDetails() AS String
		RETURN THIS.OnGetDetails()
	ENDFUNC

	* -- Show Properties code
	PROCEDURE OnShowProperties()
		DO FORM NodeProperties WITH THIS
	ENDPROC

	* override this in a subclass to display custom
	* property dialog
	PROCEDURE ShowProperties()
		THIS.OnShowProperties()
	ENDPROC
	
	FUNCTION ShowPropertiesOkay()
		RETURN !EMPTY(THIS.Options)
	ENDFUNC

	* -- Rename code
	PROCEDURE OnRenameNode()
		LOCAL lSuccess
		DO FORM NodeRename WITH THIS TO m.lSuccess
		IF VARTYPE(m.lSuccess) == 'L' AND m.lSuccess
			THIS.RefreshNode()
		ENDIF
	ENDPROC
	PROCEDURE RenameNode()
		THIS.OnRenameNode()
	ENDPROC
	

	* -- Filter code
	PROCEDURE OnShowFilter()
		DO FORM NodeFilter WITH THIS TO m.lSuccess
		IF VARTYPE(m.lSuccess) == 'L' AND m.lSuccess
			THIS.RefreshNode()
		ENDIF
	ENDPROC

	PROCEDURE ShowFilter()
		THIS.OnShowFilter()
	ENDPROC
	
	FUNCTION ShowFilterOkay()
		RETURN !THIS.EndNode
	ENDFUNC


	FUNCTION HookAddNode(oNode)
	ENDFUNC

	FUNCTION HookRemoveNode(oNode, lRemoveAll)
	ENDFUNC

	FUNCTION HookGotoNode(oNode)
	ENDFUNC

	FUNCTION HookRefreshCaption(oNode)
	ENDFUNC

	FUNCTION HookExpandNode(oNode)
	ENDFUNC


	* traverse through the tree to find the specific
	* node to expand
	FUNCTION ExpandNode(cNodeID)
		LOCAL i
		LOCAL lFound
		
		IF VARTYPE(m.cNodeID) <> 'C' OR EMPTY(m.cNodeID)
			cNodeID = THIS.NodeID
		ENDIF

		IF THIS.NodeID == m.cNodeID
			lFound = .T.
			THIS.Expanded = .T.
		ELSE
			FOR i = 1 TO THIS.oNodeList.Count
				IF THIS.oNodeList.Item(i).ExpandNode(m.cNodeID)
					EXIT
				ENDIF
			ENDFOR
		ENDIF
		
		RETURN m.lFound
	ENDFUNC

	* traverse through the tree to find the specific
	* node to collapse
	FUNCTION CollapseNode(cNodeID)
		LOCAL i
		LOCAL lFound
		
		IF VARTYPE(m.cNodeID) <> 'C' OR EMPTY(m.cNodeID)
			cNodeID = THIS.NodeID
		ENDIF

		IF THIS.NodeID == m.cNodeID
			lFound = .T.
			THIS.Expanded = .F.
		ELSE
			FOR i = 1 TO THIS.oNodeList.Count
				IF THIS.oNodeList.Item(i).CollapseNode(m.cNodeID)
					EXIT
				ENDIF
			ENDFOR
		ENDIF
		
		RETURN m.lFound
	ENDFUNC
	
	* Populate node with children
	*	[cNodeID] - node to populate
	FUNCTION Populate(cNodeID)
		LOCAL oNode
		LOCAL oRootNode
		LOCAL lSuccess
		
		lSuccess = .F.
		
		IF VARTYPE(m.cNodeID) == 'C' AND !EMPTY(m.cNodeID)
			oNode = THIS.GetNode(m.cNodeID)
		ELSE
			oNode = THIS
		ENDIF

		IF VARTYPE(m.oNode) == 'O'
			* BINDEVENT is at the root node only, so call
			* the hook event from there
			oRootNode = THIS.GetRootNode()
			IF m.oNode.oNodeList.Count > 0
				m.oRootNode.HookRemoveNode(oNode, .T.)
			ENDIF
			
			m.oNode.oNodeList.Remove(-1)

			lSuccess = m.oNode.OnPopulate()
			
			IF m.lSuccess
				IF NOT m.oNode.EndNode AND m.oNode.oNodeList.Count == 0
					THIS.AddNoChildrenNode()
				ENDIF
			ENDIF
		ENDIF
		
		RETURN m.lSuccess
	ENDFUNC

	PROCEDURE Expanded_ASSIGN(lExpanded)

		IF m.lExpanded AND !THIS.Expanded
			lExpanded = THIS.Populate()
		ENDIF

		THIS.Expanded = m.lExpanded

	ENDPROC

	* this is hooked by the TreeView control so that
	* we can tell it which node to position on
	FUNCTION GotoNode(oNode)
		LOCAL oRootNode
		
		oRootNode = THIS.GetRootNode()
		oRootNode.HookGotoNode(m.oNode)
	ENDFUNC

	* return TRUE if object name passes the filter test
	FUNCTION FilterCheck(oNode)
		LOCAL cExact
		LOCAL lFilterPass
		LOCAL oRegExpr
		LOCAL cFilterInclude
		LOCAL cFilterExclude
		
		cFilterInclude = THIS.GetOption("FilterInclude", SPACE(0))
		cFilterExclude = THIS.GetOption("FilterExclude", SPACE(0))
		
		IF EMPTY(m.cFilterInclude) AND EMPTY(m.cFilterExclude)
			RETURN .T.
		ENDIF

		cObjectName = m.oNode.NodeText

		lFilterPass = .T.
		cExact      = SET("EXACT")
		SET EXACT OFF
		
		lFilterPass = (UPPER(m.cObjectName) = UPPER(m.cFilterInclude))
		
      IF m.lFilterPass AND !EMPTY(m.cFilterExclude)
			lFilterPass = !(UPPER(m.cObjectName) = UPPER(m.cFilterExclude))
		ENDIF
		
		SET EXACT &cExact
		
		RETURN m.lFilterPass
	ENDFUNC


	* add a new child node to the current node	
	*	<oNode> = new node to add to the child node list
	FUNCTION AddNode(oNode)
		LOCAL oRootNode

		IF THIS.FilterCheck(m.oNode)
			IF THIS.oNodeList.Count > 0 AND THIS.oNodeList.Item(1).NodeID = "msg."
				THIS.oNodeList.Item(1).RemoveNode()
			ENDIF
		
			IF EMPTY(m.oNode.NodeID)
				oNode.NodeID = "user." + SYS(2015)
			ENDIF
			
			oNode.ParentID   = THIS.NodeID
			oNode.ParentNode = THIS
			oNode.NodeOrder  = THIS.oNodeList.Count
			oNode.NodeLevel  = THIS.NodeLevel + 1
			
			THIS.oNodeList.Add(m.oNode, m.oNode.NodeID)

			oRootNode = THIS.GetRootNode()
			m.oRootNode.HookAddNode(m.oNode)
		ENDIF		
	ENDFUNC

	* return the root node
	FUNCTION GetRootNode()
		LOCAL oRootNode
		
		oRootNode = THIS
		
      DO WHILE !ISNULL(m.oRootNode.ParentNode)
			oRootNode = m.oRootNode.ParentNode
		ENDDO
		
		RETURN m.oRootNode
	ENDFUNC

	* remove current node
	FUNCTION RemoveNode()
		LOCAL oRootNode
		
		THIS.Inactive = .T.


		* BINDEVENT is at the root node only, so call
		* the hook event from there
		oRootNode = THIS.GetRootNode()
		m.oRootNode.HookRemoveNode(THIS)
	ENDFUNC

	PROCEDURE ReleaseNodes()
		LOCAL i
		
		THIS.ParentNode = .NULL.
		
      IF TYPE("THIS.oNodeList") == 'O' AND !ISNULL(THIS.oNodeList)
			FOR i = THIS.oNodeList.Count TO 1 STEP -1
				THIS.oNodeList.Item(i).ReleaseNodes()
			ENDFOR
			THIS.oNodeList.Remove(-1)
		ENDIF
	ENDPROC


	* Return description of node that we can use
	* to identify node later (NodeID doesn't work
	* because that can change between sessions)
	FUNCTION GetNodeInfo()
		LOCAL cNodeInfo

		IF VARTYPE(THIS.NodeData) == 'O'
			cNodeInfo = THIS.NodeData.Type + '.' + RTRIM(THIS.NodeData.Name)
		ELSE
			cNodeInfo = RTRIM(THIS.NodeText)
		ENDIF
		RETURN m.cNodeInfo
	ENDFUNC
	
	FUNCTION RefreshNode()
		LOCAL i
		LOCAL oRootNode
	
		oRootNode = THIS.GetRootNode()
		m.oRootNode.Save()
	
		THIS.lRetrieved = .F.

		IF THIS.Expanded
			THIS.Populate()
		ENDIF

		m.oRootNode.HookRefreshCaption(THIS)
	ENDFUNC
	

	FUNCTION ExpandAll()
		LOCAL i
		
		THIS.Expanded = .T.
		
      FOR i = 1 TO THIS.oNodeList.Count
			THIS.oNodeList.Item(i).ExpandAll()
		ENDFOR
	ENDFUNC

	* Search for node by its NodeID
	FUNCTION GetNode(cNodeID)
		LOCAL i
		LOCAL oNode
		
		oNode = .NULL.
		
		IF THIS.NodeID == m.cNodeID
			oNode = THIS
		ELSE
			FOR i = 1 TO THIS.oNodeList.Count
				oNode = THIS.oNodeList.Item(i).GetNode(cNodeID)
				IF VARTYPE(m.oNode) == 'O'
					EXIT
				ENDIF
			ENDFOR
		ENDIF
		
		RETURN m.oNode
	ENDFUNC

	FUNCTION OnBeforeCreateMenu(oContextMenu)
	ENDFUNC
	FUNCTION OnAfterCreateMenu(oContextMenu)
	ENDFUNC



	FUNCTION RunScript(cUniqueID)
		LOCAL oDataExplorerEngine
		
		oDataExplorerEngine = NEWOBJECT("DataExplorerEngine", "DataExplorerEngine.prg")
		oDataExplorerEngine.RunScript(m.cUniqueID, THIS)
		oDataExplorerEngine = .NULL.
	ENDFUNC

	* Called by right-click to create a context menu
	* for this node.
	FUNCTION CreateContextMenu()
		LOCAL oContextMenu
		LOCAL nSelect
		LOCAL oCollection
		LOCAL i
		LOCAL j
		LOCAL nCnt
		LOCAL oDataExplorerEngine
		LOCAL nMenuCnt
		LOCAL ARRAY aScriptCode[1]
		
		oContextMenu = NEWOBJECT("ContextMenu", "foxmenu.prg")
		IF THIS.OnBeforeCreateMenu(m.oContextMenu)
		
			nMenuCnt = 0
			oDataExplorerEngine = NEWOBJECT("DataExplorerEngine", "DataExplorerEngine.prg")
			oCollection = m.oDataExplorerEngine.GetMenuItems(THIS)
			oDataExplorerEngine = .NULL.
			FOR m.i = 1 TO m.oCollection.Count
				IF !(m.oCollection.Item(i).Caption == "\-" AND (m.i == m.oCollection.Count OR m.i == 1))
					nMenuCnt = m.nMenuCnt + 1
					oContextMenu.AddMenu(RTRIM(m.oCollection.Item(i).Caption), "oCurrentNode.RunScript([" + m.oCollection.Item(i).UniqueID + "])")
				ENDIF
			ENDFOR

			IF !THIS.EndNode
				IF m.nMenuCnt > 0
					m.oContextMenu.AddMenu("\-")
				ENDIF
				
            * add in Refresh item
				m.oContextMenu.AddMenu(MENU_REFRESH_LOC, "oCurrentNode.RefreshNode()")
			ENDIF
			
			THIS.OnAfterCreateMenu(oContextMenu)
		ENDIF

		RETURN m.oContextMenu
	ENDFUNC
	

	FUNCTION EvalText(cScript)
		LOCAL cEvalScript
		LOCAL oException AS Exception
		
		cEvalScript = m.cScript
		IF LEFT(m.cScript, 1) == '(' AND RIGHT(m.cScript, 1) == ')'
			TRY
				cEvalScript = EVALUATE(m.cScript)
			
         CATCH TO oException
				MESSAGEBOX(m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
			ENDTRY
		ENDIF
		RETURN m.cEvalScript
	ENDFUNC

	* -- OPTION METHODS ---
	PROCEDURE Options_Assign(cOptions)
		THIS.ParseOptionString(m.cOptions)
		THIS.Options = m.cOptions
	ENDPROC

	PROCEDURE OptionData_Assign(cOptionData)
		THIS.ParseOptionDataString(m.cOptionData)
		THIS.OptionData = m.cOptionData
	ENDPROC
	
	FUNCTION OptionData_Access
		LOCAL cOptionData
		LOCAL i
		
		cOptionData = SPACE(0)
		FOR i = 1 TO THIS.oOptionCollection.Count
			IF !THIS.oOptionCollection.Item(i).OptionTemporary
				cOptionData = m.cOptionData + IIF(EMPTY(m.cOptionData), SPACE(0), CHR(13) + CHR(10)) + THIS.oOptionCollection.Item(i).OptionName + '=' + TRANSFORM(THIS.oOptionCollection.Item(i).OptionValue)
			ENDIF
		ENDFOR
		
		RETURN m.cOptionData
	ENDFUNC	

	* parse string in DataExplorer.Options, adding each option to the collection
	* (note: format is similar to an .INI file)
	FUNCTION ParseOptionString(cOptions)
		LOCAL i
		LOCAL nCnt
		LOCAL cValue
		LOCAL cAttrib
		LOCAL nPos
		LOCAL oOption
		LOCAL ARRAY aOptions[1]


		m.oOption = .NULL.
		m.nCnt = ALINES(m.aOptions, STRTRAN(m.cOptions, CHR(13), SPACE(0)), .T., CHR(10))
		FOR m.i = 1 TO m.nCnt
			IF !EMPTY(m.aOptions[m.i])
				IF LEFT(m.aOptions[m.i], 1) = '[' AND RIGHT(m.aOptions[m.i], 1) == ']'
					m.oOption = CREATEOBJECT("Option")
					m.oOption.OptionName = STREXTRACT(m.aOptions[m.i], '[', ']')

					THIS.AddOption(m.oOption)
				ELSE
					IF !ISNULL(m.oOption)
						m.nPos = AT('=', m.aOptions[m.i])
						IF m.nPos > 1
							m.cAttrib = ALLTRIM(LOWER(LEFT(m.aOptions[m.i], m.nPos - 1)))
							m.cValue  = ALLTRIM(SUBSTR(m.aOptions[m.i], m.nPos + 1))
							
							DO CASE
   							CASE m.cAttrib = "value"
   								m.oOption.OptionValue = m.cValue

   							CASE m.cAttrib = "classlib"
   								m.oOption.OptionClassLib = m.cValue

   							CASE m.cAttrib = "classname"
   								m.oOption.OptionClassName = m.cValue

   							CASE m.cAttrib = "caption"
   								m.oOption.OptionCaption = m.cValue

   							CASE m.cAttrib = "valueproperty"
   								m.oOption.ValueProperty = m.cValue

   							OTHERWISE
   								TRY
   									m.oOption.oPropertyCollection.Add(m.cValue, m.cAttrib)
   								CATCH
   								ENDTRY
							ENDCASE
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDFOR
	ENDFUNC

	* Parse option data as it is stored in OptionData
	* Look for options in the following format:
	*	ZipCode=33458
	*	Name=Ryan
	*	<cOptions> = string of options
	FUNCTION ParseOptionDataString(cOptions)
		LOCAL i
		LOCAL nCnt
		LOCAL nPos
		LOCAL ARRAY aOptions[1]

		m.nCnt = ALINES(aOptions, STRTRAN(m.cOptions, CHR(13), SPACE(0)), .T., CHR(10))
		FOR m.i = 1 TO m.nCnt
			m.nPos = AT('=', aOptions[i])
			IF m.nPos > 1
				THIS.SetOption(LEFT(aOptions[m.i], m.nPos - 1), SUBSTR(aOptions[m.i], m.nPos + 1))
			ENDIF
		ENDFOR
	ENDFUNC	
	
	* Add an option to the option collection
	FUNCTION AddOption(oOption AS Option)
		IF VARTYPE(m.oOption) == 'O' AND !EMPTY(m.oOption.OptionName)
			TRY
				THIS.oOptionCollection.Add(m.oOption, m.oOption.OptionName)
			CATCH
			ENDTRY		
		ENDIF
	ENDFUNC

	* Add an option to the option collection
	FUNCTION CreateOption(cOptionName, xDefaultValue, lTemporary)
		LOCAL oOption
		
		m.oOption = CREATEOBJECT("Option")
		m.oOption.OptionName  = m.cOptionName
		m.oOption.OptionValue = TRANSFORM(m.xDefaultValue)
		m.oOption.OptionTemporary = m.lTemporary
		
      TRY
			THIS.oOptionCollection.Add(m.oOption, m.oOption.OptionName)
		CATCH
		ENDTRY		
	ENDFUNC

	FUNCTION GetOptionByName(cOptionName)
		LOCAL oOption
		LOCAL i
		
		m.cOptionName = LOWER(m.cOptionName)
      m.oOption     = .NULL.
      
		FOR m.i = 1 TO THIS.oOptionCollection.Count
			IF LOWER(THIS.oOptionCollection.Item(m.i).OptionName) == m.cOptionName
				m.oOption = THIS.oOptionCollection.Item(m.i)
				EXIT
			ENDIF
		ENDFOR
		
		RETURN m.oOption
	ENDFUNC

	* Set an option value
	* 	<cOptionName> = option name to set
	* 	<cOptionValue> = option value
	FUNCTION SetOption(cOptionName, cOptionValue)
		LOCAL oOption
		LOCAL lSuccess
		
		m.oOption = THIS.GetOptionByName(m.cOptionName)
		m.lSuccess = VARTYPE(m.oOption) == 'O'
		IF m.lSuccess
			m.oOption.OptionValue = m.cOptionValue
		ENDIF
		
		RETURN m.lSuccess
	ENDFUNC

	* Return an option value
	* 	<cOptionName> = option name to return
	* 	[xDefault] = default to return if option not found
	*   [lSearch] = search up the nodes for option
	FUNCTION GetOption(cOptionName, xDefault, lSearch)
		LOCAL oOption
		LOCAL cOptionValue
		LOCAL cDataType
		LOCAL oNode

		IF PCOUNT() > 1
			m.cDataType = VARTYPE(m.xDefault)
		ELSE
			m.cDataType = 'C'
		ENDIF

		IF m.lSearch
			* search all parent nodes for a matching option
			oNode = THIS
			DO WHILE !ISNULL(m.oNode)
				m.oOption = m.oNode.GetOptionByName(m.cOptionName)
				IF !ISNULL(m.oOption)
					EXIT
				ENDIF
				oNode = m.oNode.ParentNode
			ENDDO
		ELSE
			m.oOption = THIS.GetOptionByName(m.cOptionName)
		ENDIF

		IF VARTYPE(m.oOption) == 'O'
			m.cOptionValue = TRANSFORM(m.oOption.OptionValue)
		ELSE
			IF PCOUNT() > 1
				m.cOptionValue = TRANSFORM(m.xDefault)
			ELSE
				m.cOptionValue = SPACE(0)
			ENDIF
		ENDIF
		
		DO CASE
		CASE m.cDataType == 'N'
			RETURN VAL(m.cOptionValue)
		CASE m.cDataType == 'L'
			RETURN (UPPER(m.cOptionValue) == ".T." OR m.cOptionValue == '1')
		OTHERWISE		
			RETURN m.cOptionValue
		ENDCASE
		
		RETURN m.xDefault
	ENDFUNC
	


	* Traverse up the tree looking for first node 
	* that has the specified option.
	FUNCTION FindOption(cOptionName, xDefaultValue)
		RETURN THIS.GetOption(m.cOptionName, m.xDefaultValue, .T.)
	ENDFUNC

	FUNCTION RemoveConnection()
		THIS.RemoveNode()
	ENDFUNC

	FUNCTION AddErrorNode(cErrorMsg)
		LOCAL oChildNode

		oChildNode = THIS.CreateNode("ErrorNode", SPACE(0), m.cErrorMsg)
		THIS.AddNode(m.oChildNode)
	ENDFUNC

	FUNCTION AddNoChildrenNode()
		LOCAL oChildNode

		oChildNode = THIS.CreateNode("NoChildrenNode")
		THIS.AddNode(m.oChildNode)
	ENDFUNC

	FUNCTION StripPassword(cStr as String)
		LOCAL cNewValue
		LOCAL i
		LOCAL cVal
		
		cNewValue = SPACE(0)
		FOR i = 1 TO GETWORDCOUNT(m.cStr, ';')
			cVal = GETWORDNUM(m.cStr, i, ';')
			IF !(UPPER(LEFT(m.cVal, 8)) = "PASSWORD" OR UPPER(LEFT(m.cVal, 3)) = "PWD")
				cNewValue = m.cNewValue + IIF(EMPTY(m.cNewValue), SPACE(0), ';') + GETWORDNUM(m.cStr, i, ';')
			ELSE
				IF AT('=', m.cVal) > 0
					cNewValue = m.cNewValue + IIF(EMPTY(m.cNewValue), SPACE(0), ';') + LEFT(m.cVal, AT('=', m.cVal)) + "***"
				ENDIF
			ENDIF
		ENDFOR
		
		RETURN m.cNewValue
	ENDFUNC


ENDDEFINE

* <summary>
* 	Root node.
* </summary>
DEFINE CLASS RootNode AS INode
	NodeID       = "root"
	NodeText     = NODETEXT_ROOT_LOC
	ExpandOnInit = .T.
	
	* Load in all active root nodes, etc
	FUNCTION OnPopulate()
		LOCAL oChildNode
		LOCAL oDataExplorerEngine
		LOCAL oCollection
		LOCAL nIndex
		
		oDataExplorerEngine = NEWOBJECT("DataExplorerEngine", "DataExplorerEngine.prg")

		oCollection = m.oDataExplorerEngine.GetRootNodes()
		
      FOR nIndex = 1 TO m.oCollection.Count
			oChildNode            = THIS.CreateNode(m.oCollection.Item(nIndex).ClassName, m.oCollection.Item(nIndex).ClassLib)
			oChildNode.Options    = m.oCollection.Item(nIndex).Options
			oChildNode.OptionData = m.oCollection.Item(nIndex).OptionData

			THIS.AddNode(m.oChildNode)
		ENDFOR
		
		oDataExplorerEngine = .NULL.
		
		RETURN .T.
	ENDFUNC

	* Persist all user-added connections in the tree
	FUNCTION Save()
		LOCAL oDataExplorerEngine 

		m.oDataExplorerEngine = NEWOBJECT("DataExplorerEngine", "DataExplorerEngine.prg")
		m.oDataExplorerEngine.Save(THIS)
		m.oDataExplorerEngine = .NULL.
	ENDFUNC

	FUNCTION ShowFilterOkay()
		RETURN .F.
	ENDFUNC

ENDDEFINE


* <summary>
* 	Connections node.
* </summary>
DEFINE CLASS ConnectionsNode AS INode
	ImageKey = "microsoft.imageconns"
	NodeID   = "connections"
	NodeText = NODETEXT_CONNECTIONS_LOC
	
	ExpandOnInit = .T.
	
	* Load in all active connections, etc
	FUNCTION OnPopulate()
		LOCAL oChildNode
		LOCAL oDataExplorerEngine
		LOCAL oCollection
		LOCAL i
		
		
		oDataExplorerEngine = NEWOBJECT("DataExplorerEngine", "DataExplorerEngine.prg")
		oCollection         = m.oDataExplorerEngine.GetActiveConnections()
		
      FOR i = 1 TO m.oCollection.Count
			THIS.AddConnectionNode(m.oCollection.Item(i))
		ENDFOR
		
      oDataExplorerEngine = .NULL.
		
		RETURN .T.
	ENDFUNC




	* check to see if connection is is a duplicate and
	* if so ask to rename it.
	FUNCTION CheckDuplicateName(oNode)
		LOCAL i
		LOCAL oConnCollection
		LOCAL oDataExplorerEngine
		LOCAL lSuccess
		LOCAL oRootNode
		
		lSuccess = .T.

		oRootNode = THIS.GetRootNode()
		m.oRootNode.Save()
		
		* if this is a duplicate, then automatically display rename dialog
		oDataExplorerEngine = NEWOBJECT("DataExplorerEngine", "DataExplorerEngine.prg")
		oConnCollection     = m.oDataExplorerEngine.GetActiveConnections()
		
      FOR i = 1 TO m.oConnCollection.Count
			IF UPPER(m.oConnCollection.Item(i).ConnName) == UPPER(m.oNode.NodeText)
				DO FORM NodeRename WITH oNode, .T. TO m.lSuccess
				EXIT
			ENDIF
		ENDFOR
      
		oDataExplorerEngine = .NULL.
		
		RETURN m.lSuccess
	ENDFUNC
	
	FUNCTION AddConnectionNode(oConn, lUserAdded)
		LOCAL nCnt
		LOCAL nPropIndex
		LOCAL cPropertyName
		LOCAL oNode
		LOCAL oException AS Exception 

		TRY
			oNode = THIS.CreateNode(m.oConn.ClassName, m.oConn.ClassLib)

			oNode.NodeID     = m.oConn.UniqueID
			oNode.DefType    = DEFTYPE_CONNECTION
			oNode.Options    = m.oConn.Options
			oNode.OptionData = m.oConn.OptionData

			IF NOT EMPTY(m.oConn.DataMgmtClass) AND NOT EMPTY(m.oConn.DataMgmtClassLibrary)
				oNode.DataMgmtClass        = m.oConn.DataMgmtClass
				oNode.DataMgmtClassLibrary = m.oConn.DataMgmtClassLibrary
			ENDIF
			
         oNode.ProviderName = m.oConn.ProviderName

			* if this connection is being added for the first time, then
			* ask the user for connection info in the OnFirstConnect() method
			IF NOT m.lUserAdded OR m.oNode.FirstConnect(m.oConn)
				oNode.NodeText = EVL(m.oConn.ConnName, m.oNode.NodeText)
				oNode.NodeType = m.oConn.ConnType
				*oNode.NodeInfo = oConn.ConnInfo
				
				oNode.SaveNode = .T.

				IF NOT m.lUserAdded OR THIS.CheckDuplicateName(m.oNode)
					oConn.ConnName = m.oNode.NodeText
					THIS.AddNode(m.oNode)
					IF m.lUserAdded
						THIS.GotoNode(m.oNode)
					ENDIF
				ENDIF
			ENDIF
		CATCH TO m.oException
			MESSAGEBOX(ERROR_CREATECONNECTION_LOC + CHR(10) + CHR(10) + m.oException.Message, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
		ENDTRY

	ENDFUNC
	

	* Add a new connection	
	FUNCTION AddConnection(cConnTypeUniqueID)
		LOCAL oConn
		LOCAL oDataExplorerEngine
		
		oConn = .NULL.

		IF VARTYPE(m.cConnTypeUniqueID) == 'C' AND NOT EMPTY(m.cConnTypeUniqueID)
			oDataExplorerEngine = NEWOBJECT("DataExplorerEngine", "DataExplorerEngine.prg")
			oConn = m.oDataExplorerEngine.CreateConnection(m.cConnTypeUniqueID)
			oDataExplorerEngine = .NULL.
		ENDIF
		
		IF VARTYPE(m.oConn) <> 'O'
			DO FORM AddConnection TO m.oConn
		ENDIF
		
		IF VARTYPE(m.oConn) == 'O'
			* we now have the type of connection we want to add
			* (VFP Database, VFP Table, SQL Server, SQL Database, etc),
			* so now get additional info depending on the
			* connection type
			THIS.AddConnectionNode(m.oConn, .T.)
		ENDIF
		
	ENDFUNC

	FUNCTION ShowFilterOkay()
		RETURN .F.
	ENDFUNC

ENDDEFINE


* Interface for connection nodes
DEFINE CLASS IConnectionNode AS INode
	* Data managament class library and class to use
	DataMgmtClassLibrary = SPACE(0)
	DataMgmtClass        = SPACE(0)
	ProviderName         = SPACE(0)

	* if we run the Script Code when we first create the connection and our
	* Data Mgmt Class changes, then set this to TRUE
	CustomDataMgmtClass = .F. 

	PROCEDURE OnInit()
		DODEFAULT()

		THIS.CreateOption("ShowColumnInfo", .F.)
		THIS.CreateOption("TrustedConnection", .T.)
		THIS.CreateOption("UserName", SPACE(0))
	ENDPROC

	PROCEDURE FirstConnect(oConn)
		LOCAL lSuccess
		LOCAL oException AS Exception
		LOCAL cDataMgmtClassLibrary
		LOCAL cDataMgmtClass
		LOCAL oTHIS
		
		lSuccess = THIS.OnFirstConnect()
		
      IF lSuccess
			* if connection object has ScriptCode defined, then
			* we execute that now so properties can be changed, etc
			* based upon the connection created
			IF !EMPTY(m.oConn.ScriptCode)
				TRY 
					cDataMgmtClassLibrary = UPPER(THIS.DataMgmtClassLibrary)
					cDataMgmtClass        = UPPER(THIS.DataMgmtClass)

					oTHIS = THIS
					lSuccess = EXECSCRIPT(m.oConn.ScriptCode, m.oTHIS, THIS.GetConnection())
					IF VARTYPE(lSuccess) <> 'L'  && if script returns anything but T/F, then assume true
						lSuccess = .T.
					ENDIF
						
					THIS.CustomDataMgmtClass = NOT (UPPER(THIS.DataMgmtClassLibrary) == m.cDataMgmtClassLibrary) OR !(UPPER(THIS.DataMgmtClass) == m.cDataMgmtClass)
				
            CATCH TO m.oException
					MESSAGEBOX(ERROR_CREATECONNSCRIPT_LOC + CHR(10) + CHR(10) + m.oException.Message + CHR(10) + m.oException.LineContents, MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
				ENDTRY
			ENDIF
		ENDIF
		
		RETURN m.lSuccess
	ENDPROC
	
	* called the first time a node of this type is created
	* (not instantiated -- but created.  For example, the user
	* right-clicks on Connections and selects "Create New Connection"
	FUNCTION OnFirstConnect()
		THIS.SetOption("ShowColumnInfo", THIS.GetConfigValue("ShowColumnInfo", .F.))
		RETURN .T.
	ENDFUNC

	FUNCTION GetConnection(cServerName, cDatabaseName)
		RETURN .NULL.
	ENDFUNC

	
	FUNCTION GetDataMgmtObject()
		LOCAL oException AS Exception
		LOCAL oDataMgmt
		
		oDataMgmt = .NULL.
      
		TRY
			oDataMgmt = NEWOBJECT(THIS.DataMgmtClass, THIS.DataMgmtClassLibrary)
		CATCH TO m.oException
			MESSAGEBOX(ERROR_CREATEDATAMGMT_LOC + CHR(10) + CHR(10) + m.oException.Message + CHR(10) + CHR(10) + "(" + THIS.DataMgmtClass + "," + THIS.DataMgmtClassLibrary + ")", MB_ICONSTOP, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
		ENDTRY
		
      RETURN m.oDataMgmt
	ENDFUNC

	* return a list of available SQL servers
	FUNCTION GetAvailableServers()
		LOCAL oConn
		LOCAL oServerList

		* create a SQL database management object, which we'll
		* use to get available SQL servers
		oConn = THIS.GetConnection()
		oServerList = oConn.GetAvailableServers()
		
		RETURN m.oServerList		
	ENDFUNC

	* return a list of available database for given server	
	FUNCTION GetAvailableDatabases(cServerName)
		LOCAL oConn
		LOCAL oDatabaseList
		LOCAL i
		LOCAL oException AS Exception 
		
		oDatabaseList = .NULL.		
		
      TRY
			oConn = THIS.GetDataMgmtObject()
			
         IF oConn.Connect(m.cServerName)
				oDatabaseList = m.oConn.GetDatabases()
			ENDIF
		
      CATCH TO oException
			MESSAGEBOX(IIF(VARTYPE(m.oException.UserValue) == 'C' AND !EMPTY(m.oException.UserValue), m.oException.UserValue, m.oException.Message), MB_ICONEXCLAMATION, DATAEXPLORER_LOC + " - " + LOWER(m.oException.Procedure))
		ENDTRY
		
		RETURN m.oDatabaseList
	ENDFUNC


	FUNCTION DataTypeToString(cDataType, nLength, nDecimals)
		LOCAL cDisplayAs
		
		cDisplayAs = m.cDataType
		IF VARTYPE(m.nLength) == 'N' AND nLength > 0
			cDisplayAs = m.cDisplayAs + " (" + TRANSFORM(m.nLength) + IIF(VARTYPE(m.nDecimals) == 'N', ", " + TRANSFORM(m.nDecimals), SPACE(0)) + ")"
		ENDIF
		
		RETURN m.cDisplayAs
	ENDFUNC

	FUNCTION ExecuteQuery(cSQL, cAlias) AS String
		LOCAL oConn
		LOCAL cResult

		* hand off to the connection object
		oConn = THIS.GetConnection()
		oConn.ExecuteQuery(m.cSQL, m.cAlias)

		cResult = m.oConn.QueryResultOutput
		
		oConn = .NULL.
		
		RETURN m.cResult
	ENDFUNC


	FUNCTION RunQuery(cSQL)
		LOCAL oConn

		* hand off to the connection object
		IF VARTYPE(m.cSQL) <> 'C'
			cSQL = THIS.OnDefaultQuery()
		ENDIF

		IF VARTYPE(m.cSQL) == 'C'
			oConn = THIS.GetConnection()
			oConn.RunQuery(m.cSQL)
		ENDIF
	ENDFUNC


	* return default query for Run Query statement for this specific node type
	FUNCTION OnDefaultQuery()
		RETURN SPACE(0)
	ENDFUNC


	* delimit passed object name if spaces found
	FUNCTION FixName(cObjName)
		IF SPACE(1) $ m.cObjName
			RETURN '"' + m.cObjName + '"'
		ELSE
			RETURN m.cObjName
		ENDIF
	ENDFUNC

ENDDEFINE


DEFINE CLASS ErrorNode AS INode OF TreeNodes.prg
	ImageKey = "microsoft.imageerror"
	EndNode  = .T.

	PROCEDURE Init(cNodeText, xNodeData)
		THIS.NodeID = "msg." + SYS(2015)

		DODEFAULT(m.cNodeText, m.xNodeData)
	ENDPROC

ENDDEFINE

DEFINE CLASS NotSupportedNode AS INode OF TreeNodes.prg
	ImageKey = "microsoft.imageerror"
	EndNode  = .T.

	NodeText = NOT_SUPPORTED_LOC
	
	PROCEDURE Init(cNodeText, xNodeData)
		THIS.NodeID = "msg." + SYS(2015)

		DODEFAULT(m.cNodeText, m.xNodeData)
	ENDPROC

ENDDEFINE

DEFINE CLASS NoChildrenNode AS INode OF TreeNodes.prg
	ImageKey = ""
	EndNode  = .T.
	
	NodeText = NO_CHILDREN_LOC
	
	PROCEDURE Init(cNodeText, xNodeData)
		THIS.NodeID = "msg." + SYS(2015)

		DODEFAULT(m.cNodeText, m.xNodeData)
	ENDPROC
ENDDEFINE


DEFINE CLASS Option AS Custom
	ADD OBJECT oPropertyCollection AS Collection

	OptionName        = SPACE(0)
	OptionValue       = SPACE(0)
	OptionTemporary   = .F.
	OptionCaption     = SPACE(0)
	OptionClassName   = "cfoxtextbox"
	OptionClassLib    = SPACE(0)
	ValueProperty     = "value"
ENDDEFINE

*: EOF :*
