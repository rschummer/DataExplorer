* Constants related to errors.

* Error resolution return values.

#define ccMSG_RETRY					'Retry'
#define ccMSG_CONTINUE				'Continue'
#define ccMSG_CLOSEFORM				'Close Form'
#define ccMSG_QUIT					'Quit'
#define ccMSG_CANCEL				'Cancel'
#define ccMSG_DEBUG                 'Debug'

* Error message strings.

#define ccMSG_ERROR_NUM				'Error #:'
#define ccMSG_MESSAGE				'Message:'
#define ccMSG_LINE_NUM				'Line #:'
#define ccMSG_CODE					'Code:'
#define ccMSG_METHOD				'Method:'

* AERROR() array dimensions, including added extensions:

#define cnAERR_NUMBER				 1
#define cnAERR_MESSAGE				 2
#define cnAERR_OBJECT				 3
#define cnAERR_WORKAREA				 4
#define cnAERR_TRIGGER				 5
#define cnAERR_EXTRA1				 6
#define cnAERR_EXTRA2				 7
#define cnAERR_METHOD				 8
#define cnAERR_LINE					 9
#define cnAERR_SOURCE				10
#define cnAERR_DATETIME				11
#define cnAERR_USER					12
#define cnAERR_MAX					cnAERR_USER

* VFP error numbers:

#define cnERR_FILE_NOT_FOUND		   1
	* File does not exist
#define cnERR_ARGUMENT_INVALID        11
	* Function argument value, type, or count is invalid
#define cnERR_ALIAS_NOTFOUND		  13
	* Alias is not found
#define cnERR_INDEX_FILE_TABLE        19
	* Index file doesn't match table
#define cnERR_INVALID_SUBSCRIPT		  31
	* Invalid subscrip reference
#define cnERR_NO_FIELDS               47
	* No fields found to process
#define cnERR_NO_TABLE_SELECTED       52
	* No table is open in the current work area
#define cnERR_FILEINUSE				 108
	* File in use by another user
#define cnERR_RECINUSE				 109
	* Record in use by another user
#define cnERR_INDEX_MATCH_TABLE      114
	* Index doesn't match table
#define cnERR_ARRAYDIM				 230
   * Array dimensions are invalid
#define cnERR_USER_DEFINED			1098
	* User-defined error
#define cnERR_TOO_FEW_ARGS			1229
	* Too few arguments
#define cnERR_SUBSCRIPT_RANGE       1234
	* Subscript is outside defined range
#define cnERR_OLE_EXEC_FAIL         1426
	* OLE error: remote procedure call failed
#define cnERR_OLE_ERROR             1429
	* OLE error
#define cnERR_EXECUTION_CANCELED    1523
	* Execution was canceled by user
#define cnERR_TRIGGER_FAILED		1539
	* Trigger failed
#define cnERR_PROPERTY_INVALID      1560
	* Property value invalid
#define cnERR_PRIM_KEY_INVALID		1567
	* Primary key property invalid
#define cnERR_NONULLS				1581
	* Field does not accept null values
#define cnERR_FIELD_RULE_FAILED		1582
	* Field validation rule is violated
#define cnERR_TABLE_RULE_FAILED		1583
	* Record validation rule is violated
#define cnERR_RECMODIFIED			1585
	* Update conflict
#define cnERR_TAG_NOT_FOUND         1683
	* Index tag is not found
#define cnERR_ACCESS_DENIED			1705
	* File access is denied
#define cnERR_CDX_NOT_FOUND			1707
	* Structural CDX not found
#define cnERR_PROPERTY_TYPE_INVALID 1732
	* Data type invalid for this property
#define cnERR_PROPERTY_READ_ONLY	1743
	* Property is read-only
#define cnERR_DUPLKEY				1884
	* Uniqueness of index is violated
#define cnERR_DE_UNLOADED			1967
	* Data environment is already unloaded
#define cnERR_TABLE_IN_USE			1995
	* Error loading the data environment: table is in use
#define cnERR_TABLE_MOVED			2004
	* The table has moved
#define cnERR_CANT_SET_FOCUS        2012
	* Cannot call SetFocus from within a When, Valid ...
