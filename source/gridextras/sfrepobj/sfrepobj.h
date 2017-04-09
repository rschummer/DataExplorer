* Totaling types.

#define ccTOTAL_NONE                'N'
#define ccTOTAL_COUNT               'C'
#define ccTOTAL_SUM                 'S'
#define ccTOTAL_AVERAGE             'A'
#define ccTOTAL_LOWEST              'L'
#define ccTOTAL_HIGHEST             'H'
#define ccTOTAL_STDDEV              'D'
#define ccTOTAL_VARIANCE            'V'

* Object alignment constants.

#define ccALIGN_LEFT                'Left'
#define ccALIGN_CENTER              'Center'
#define ccALIGN_RIGHT               'Right'
#define cnALIGN_LEFT                0
#define cnALIGN_RIGHT               1
#define cnALIGN_CENTER              2

* Report font style constants.

#define cnSTYLE_NORMAL              0
#define cnSTYLE_BOLD                1
#define cnSTYLE_ITALIC              2
#define cnSTYLE_UNDERLINE           4

* Report measurement units.

#define ccUNITS_CHARACTERS          'C'
#define ccUNITS_INCHES              'I'
#define ccUNITS_CENTIMETERS         'M'

* Report object types.

#define cnREPOBJ_HEADER             1
#define cnREPOBJ_TEXT               5
#define cnREPOBJ_FIELD              8
#define cnREPOBJ_BAND               9

* Report band names (in upper case).

#define ccBAND_PAGE_HEADER          'PAGE HEADER'
#define ccBAND_PAGE_FOOTER          'PAGE FOOTER'
#define ccBAND_DETAIL               'DETAIL'
#define ccBAND_TITLE                'TITLE'
#define ccBAND_SUMMARY              'SUMMARY'
#define ccBAND_GROUP_HEADER         'GROUP HEADER'
#define ccBAND_GROUP_FOOTER         'GROUP FOOTER'

* Report object types (in upper case).

#define ccOBJECT_FIELD              'FIELD'
#define ccOBJECT_TEXT               'TEXT'
#define ccOBJECT_LINE               'LINE'
#define ccOBJECT_IMAGE              'IMAGE'
#define ccOBJECT_BOX                'BOX'

* Object positions options.

#define cnPOSITION_FLOAT            0
#define cnPOSITION_TOP              1
#define cnPOSITION_BOTTOM           2

* Other constants.

#define cnPOINTS_PER_INCH           72

* cnVERTICAL_FUDGE is a "fudge factor" used so objects are correctly placed.

#define cnVERTICAL_FUDGE            2

* cnPEN_RED_FUDGE is a "fudge factor" for the PEN_RED value in an FRX.

#define cnPEN_RED_FUDGE             4

* cnREPORT_UNITS is the number of report units (a value used for all
* measurements inside the FRX) per inch

#define cnREPORT_UNITS              10000

* cnFACTOR is the number of report units per pixel. It's calculated as report
* units per inch (10,000) divided by pixels per inch (96). 10000/96 = 104.166.

#define cnFACTOR                    104.166            

* cnBAND_HEIGHT is the height of the band separator, which is 20 pixels or
* 2083.333 report units.

#define cnBAND_HEIGHT               2083.333

* cnGROUP_OFFSET is what to add to a group number for group breaks.

#define cnGROUP_OFFSET              6

* ccAVG_CHAR is the "average" character in a font used to determine object
* sizing.

#define ccAVG_CHAR                  'N'

* Error numbers.

#define cnERR_USER_DEFINED			1098
	* User-defined error
#define cnERR_PROPERTY_INVALID      1560
	* Property value invalid
#define cnERR_PROPERTY_TYPE_INVALID 1732
	* Data type invalid for this property

* Error messages.

#define ccERR_NO_REPORT_FILE        'The name of the report file was not specified'
