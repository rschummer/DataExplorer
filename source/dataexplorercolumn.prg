* RAS 3-Jul-2011, made pass for m.Variables

* Defines the column class to use in for the grid in the RunQuery form
DEFINE CLASS CQueryColumn AS Column

	PROCEDURE Init()
		THIS.FontName   = THIS.Parent.FontName
		THIS.FontSize   = THIS.Parent.FontSize
		THIS.FontBold   = THIS.Parent.FontBold
		THIS.FontItalic = THIS.Parent.FontItalic
		THIS.ReadOnly   = THIS.Parent.ReadOnly
	ENDPROC
	
*!*      * RAS 9-Oct-2006, added code to have the header font and attributes match column
*!*      ADD OBJECT CQueryHeader AS CQueryHeader WITH ;
*!*       Name = "grcHeader", ;
*!*       FontName = THIS.FontName, ;
*!*       FontSize = THIS.FontSize, ;
*!*       FontBold = THIS.FontBold, ;
*!*       FontItalic = THIS.FontItalic

	ADD OBJECT CQueryTextbox AS CQueryTextBox WITH ;
      Name       = "QueryTextbox", ;
      FontName   = THIS.FontName, ;
      FontSize   = THIS.FontSize, ;
      FontBold   = THIS.FontBold, ;
      FontItalic = THIS.FontItalic
	
   CurrentControl = "QueryTextbox"
ENDDEFINE


DEFINE CLASS CQueryHeader AS Header

   PROCEDURE Init

   ENDPROC

   PROCEDURE Destroy

   ENDPROC

ENDDEFINE



DEFINE CLASS CQueryTextbox AS TextBox
	PROTECTED PROCEDURE ShowBrowseWindow()
      * RAS 9-Oct-2006, added code to deal with general fields not previously handled
      LOCAL lcType

      DO CASE 
         CASE LOWER(TRANSFORM(this.Value)) = "gen"
            * Deal with General field
            lcType = "General"      
         OTHERWISE 
            lcType = "MemoBlob"
      ENDCASE
   
		IF THIS.Parent.ReadOnly
         IF lcType = "General"
            MODIFY GENERAL (THIS.ControlSource) NOMODIFY
         ELSE
			   MODIFY MEMO (THIS.ControlSource) IN MACDESKTOP NOEDIT
         ENDIF
		ELSE
         IF lcType = "General"
            MODIFY GENERAL (THIS.ControlSource) 
         ELSE
			   MODIFY MEMO (THIS.ControlSource) IN WINDOW (THISFORM.Name)
         ENDIF
		ENDIF
	ENDPROC

	PROCEDURE DblClick()
		IF INLIST(TYPE(THIS.ControlSource), 'M', 'G', 'W')
			THIS.ShowBrowseWindow()
			NODEFAULT
		ENDIF
	ENDPROC

	PROCEDURE KeyPress(nKeyCode, nShiftAltCtrl)
		* ctrl+pgdn or ctrl+home
		IF ((nKeyCode == 30 AND nShiftAltCtrl == 2) OR (nKeyCode == 29 AND nShiftAltCtrl == 2)) AND ;
		 INLIST(TYPE(THIS.ControlSource), 'M', 'G', 'W')
			THIS.ShowBrowseWindow()
			NODEFAULT
		ENDIF
	ENDPROC

ENDDEFINE

*: EOF :*