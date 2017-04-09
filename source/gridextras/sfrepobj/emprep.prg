* This test uses "characters" for units.

if used('EMPLOYEE')
	select EMPLOYEE
else
	select 0
	use (_samples + 'DATA\EMPLOYEE')
endif used('EMPLOYEE')
set database to TESTDATA

* Create an SFReportFile object and set some properties.

loReport = newobject('SFReportFile', 'SFRepObj')
loReport.cReportFile  = 'EMPREP.FRX'
loReport.lSummaryBand = .T.
loReport.cFontName    = 'Arial'
loReport.cUnits       = 'Inches'
loReport.nLeftMargin  = 0.5

* Set the height of the page header band.

loPageHeader = loReport.GetReportBand('Page Header')
loPageHeader.nHeight = 1

* Set the height of the detail band.

loDetail = loReport.GetReportBand('Detail')
loDetail.nHeight = 3

* Set the height of the summary band.

loSummary = loReport.GetReportBand('Summary')
loSummary.nHeight = 0.5

* Insert column heading and field records into the appropriate band.
	
lnFields = afields(laFields)
for lnI = lnFields to 1 step -1
	if not inlist(laFields[lnI, 1], 'FIRST_NAME', 'LAST_NAME', 'PHOTO')
		adel(laFields, lnI)
		dimension laFields[alen(laFields, 1) - 1, alen(laFields, 2)]
	endif not inlist(laFields[lnI, 1], ...
next lnI
lnFields = alen(laFields, 1)
lnVPos   = 0
lnWidth  = 0
lnSeparation = 0.02
for lnI = 1 to lnFields

* Get the attributes for the current column.

	lcField  = laFields[lnI, 1]
	lcAField = 'EMPLOYEE.' + lcField
	lcColumn = dbgetprop(lcAField, 'Field', 'Caption')
	if empty(lcColumn)
		lcColumn = proper(strtran(lcField, '_', ' '))
	endif empty(lcColumn)
	lcColumn = lcColumn + ':'

* Insert the field heading.

	loObject = loDetail.AddItem('Text')
	loObject.cExpression = lcColumn
	loObject.nVPosition  = lnVPos
	loObject.nHPosition  = 0
	loObject.lFontBold   = .T.

* Insert the field.

	if laFields[lnI, 2] = 'G'
		loObject = loDetail.AddItem('Image')
		loObject.cImageSource = lcAField
		loObject.lImageFile   = .F.
		loObject.nWidth       = 2
		loObject.nHeight      = 2
		loObject.nVPosition   = lnVPos + 0.05
		loObject.nHPosition   = 1.05
		loObject = loDetail.AddItem('Box')
		loObject.nWidth     = 2.1
		loObject.nHeight    = 2.1
		loObject.nVPosition = lnVPos
		loObject.nHPosition = 1
	else
		loObject = loDetail.AddItem('Field')
		loObject.cExpression = lcAField
		loObject.nWidth      = 5
		loObject.nVPosition  = lnVPos
		loObject.nHPosition  = 1
	endif laFields[lnI, 2] = 'G'

* Update the current vertical position and report width.

	lnWidth = max(lnWidth, loObject.nHPosition + loObject.nWidth)
	lnVPos  = lnVPos + loObject.nHeight + lnSeparation
next lnI

* Insert a line at the bottom of the detail band.

loObject = loDetail.AddItem('Line')
loObject.nWidth     = lnWidth
loObject.nVPosition = lnVPos + 0.3
loObject.nHPosition = 0

* Insert the page header.

loObject = loPageHeader.AddItem('Field')
loObject.cExpression = '"Employee List"'
loObject.nWidth      = lnWidth
loObject.nHPosition  = 0
loObject.nVPosition  = 0
loObject.cAlignment  = 'Center'
loObject.lFontBold   = .T.
loObject.nFontSize   = 24
loObject.nForeColor  = rgb(255, 0, 0)

* Insert the date.

loObject = loPageHeader.AddItem('Field')
loObject.cExpression = 'date()'
loObject.nWidth      = 1
loObject.nHPosition  = 0
loObject.nVPosition  = 0.5
loObject.lFontBold   = .T.

* Insert the page number.

lnPagenoWidth = 0.2
loObject = loPageHeader.AddItem('Text')
loObject.cExpression = 'Page'
loObject.nVPosition  = 0.5
loObject.nHPosition  = lnWidth - lnPagenoWidth - 0.4
loObject.lFontBold   = .T.
loObject = loPageHeader.AddItem('Field')
loObject.cExpression = '_pageno'
loObject.nWidth      = lnPagenoWidth
loObject.nVPosition  = 0.5
loObject.nHPosition  = lnWidth - lnPagenoWidth
loObject.cAlignment  = 'Right'
loObject.lFontBold   = .T.

* Insert a line.

loObject = loPageHeader.AddItem('Line')
loObject.nWidth      = lnWidth
loObject.nVPosition  = 0.75
loObject.nHPosition  = 0
loObject.nPenSize    = 6

* Add a total in the summary band.

loObject = loSummary.AddItem('Field')
loObject.cExpression = 'ltrim(str(lnCount)) + " record" + ' + ;
	'iif(lnCount = 1, "", "s") + " printed"'
loObject.nWidth      = 2
loObject.nVPosition  = 0.1
loObject.nHPosition  = 0
loObject.lFontBold   = .T.

* Create the report variables.

loVariable = loReport.CreateVariable()
loVariable.cName         = 'lnCount'
loVariable.cValue        = 1
loVariable.cInitialValue = 0
loVariable.cTotalType    = 'Sum'

* Set the height of the page footer band.

loPageFooter = loReport.GetReportBand('Page Footer')
loPageFooter.nHeight = 0

* Create the report, then let's run it.

loReport.Save()
report form (loReport.cReportFile) preview

* Specifically release the report object so proper object cleanup occurs.

loReport.Release()
