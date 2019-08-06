/****************************************************************************
** Copyright (c) 2013-2014 Debao Zhang <hello@debao.me>
** All right reserved.
**
** Permission is hereby granted, free of charge, to any person obtaining
** a copy of this software and associated documentation files (the
** "Software"), to deal in the Software without restriction, including
** without limitation the rights to use, copy, modify, merge, publish,
** distribute, sublicense, and/or sell copies of the Software, and to
** permit persons to whom the Software is furnished to do so, subject to
** the following conditions:
**
** The above copyright notice and this permission notice shall be
** included in all copies or substantial portions of the Software.
**
** THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
** EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
** MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
** NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
** LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
** OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
** WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
**
****************************************************************************/
#include "xlsxworkbook.h"
#include "xlsxworkbook_p.h"
#include "xlsxsharedstrings_p.h"
#include "xlsxworksheet.h"
#include "xlsxchartsheet.h"
#include "xlsxstyles_p.h"
#include "xlsxformat.h"
#include "xlsxworksheet_p.h"
#include "xlsxformat_p.h"
#include "xlsxmediafile_p.h"
#include "xlsxutility_p.h"
#include "xlsxcellrange.h"		//	Dave added
#include "xlsxdefinednames.h"	//	Dave added

#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QFile>
#include <QBuffer>
#include <QDir>
#include <QList>

QT_BEGIN_NAMESPACE_XLSX

WorkbookPrivate::WorkbookPrivate(Workbook *q, Workbook::CreateFlag flag) :
    AbstractOOXmlFilePrivate(q, flag)
{
    sharedStrings = QSharedPointer<SharedStrings> (new SharedStrings(flag));
    styles = QSharedPointer<Styles>(new Styles(flag));
    theme = QSharedPointer<Theme>(new Theme(flag));
	definedNames.reset( new DefinedNames(q));

/*
    x_window = 240;
    y_window = 15;
    window_width = 16095;
    window_height = 9660;
	activesheetIndex = 0;
	firstsheet = 0;
	*/
    strings_to_numbers_enabled = false;
    strings_to_hyperlinks_enabled = true;
    html_to_richstring_enabled = false;
    //date1904 = false;
    defaultDateFormat = QStringLiteral("yyyy-mm-dd");
    table_count = 0;

    last_worksheet_index = 0;
    last_chartsheet_index = 0;
    last_sheet_id = 0;

	if( flag == Workbook::CreateFlag::F_NewFromScratch )
	{
		initalizeBlank();
	}

}

void WorkbookPrivate::initalizeBlank()
{
	//	Initialize for blank sheet
	//	Initialize bare minumum for fileVersion
	fileVersion.insert( QStringLiteral( "appName" ), QStringLiteral( "xl" ) );
	fileVersion.insert( QStringLiteral( "lastEdited" ), QStringLiteral( "4" ) );
	fileVersion.insert( QStringLiteral( "lowestEdited" ), QStringLiteral( "4" ) );
	fileVersion.insert( QStringLiteral( "appName" ), QStringLiteral( "4505" ) );

	//	Initialize bare minumum for fileVersion
	workbookPr.insert( QStringLiteral( "defaultThemeVersion" ), QStringLiteral( "124226" ) );

	//	Initialize bare minumum for calcPr
	calcPr.insert( QStringLiteral( "calcId" ), QStringLiteral( "124519" ) );
	calcPr.insert( QStringLiteral( "fullCalcOnLoad" ), QStringLiteral( "1" ) );

	//	Initialize workbookView
	workbookView.insert( QStringLiteral( "xWindow" ), QStringLiteral( "240" ) );
	workbookView.insert( QStringLiteral( "yWindow" ), QStringLiteral( "15" ) );
	workbookView.insert( QStringLiteral( "windowWidth" ), QStringLiteral( "16095" ) );
	workbookView.insert( QStringLiteral( "windowHeight" ), QStringLiteral( "9660" ) );
	
	//	Not needed since 0 is not saved
	//activesheetIndex = 0;
	//firstsheet = 0;
}

Workbook::Workbook(CreateFlag flag)
    : AbstractOOXmlFile(new WorkbookPrivate(this, flag))
{
}

Workbook::~Workbook()
{
}

bool Workbook::isDate1904() const
{
    Q_D(const Workbook);
	
	//	Custom class converts between 1/0 for true/false
	return d->workbookPr.toBool( QStringLiteral( "date1904" ) );
}

/*!
  Excel for Windows uses a default epoch of 1900 and Excel
  for Mac uses an epoch of 1904. However, Excel on either
  platform will convert automatically between one system
  and the other. Qt Xlsx stores dates in the 1900 format
  by default.

  \note This function should be called before any date/time
  has been written.
*/
void Workbook::setDate1904(bool date1904)
{
    Q_D(Workbook);

	if( date1904 )
		d->workbookPr.setBool( QStringLiteral( "date1904" ), date1904 );
	else
		d->workbookPr.remove( QStringLiteral( "date1904" ) );
}

//	Encapsulate active sheet logic since it is saved in several places.
//	If saved incorrectly, multiple tabs will be selected 
qint32 Workbook::activeSheetId() const
{
	Q_D(const Workbook);

	//	Custom class converts between 1/0 for true/false
	return d->workbookView.toInt( QStringLiteral( "activeTab" ) );
}

void Workbook::setActiveSheetId(qint32 id)
{
	Q_D(Workbook);

	//	Disable active tab on old sheet
	Worksheet *sheet = dynamic_cast<Worksheet*>(d->sheets[ activeSheetId() ].data());

	if( sheet )
		sheet->setSelected( false );
	
	sheet = dynamic_cast<Worksheet*>(d->sheets[ id ].data());

	if( sheet )
		sheet->setSelected( true );

	d->workbookView.setInt( QStringLiteral( "activeTab" ), id );
}

/*
  Enable the worksheet.write() method to convert strings
  to numbers, where possible, using float() in order to avoid
  an Excel warning about "Numbers Stored as Text".

  The default is false
 */
void Workbook::setStringsToNumbersEnabled(bool enable)
{
    Q_D(Workbook);
    d->strings_to_numbers_enabled = enable;
}

bool Workbook::isStringsToNumbersEnabled() const
{
    Q_D(const Workbook);
    return d->strings_to_numbers_enabled;
}

void Workbook::setStringsToHyperlinksEnabled(bool enable)
{
    Q_D(Workbook);
    d->strings_to_hyperlinks_enabled = enable;
}

bool Workbook::isStringsToHyperlinksEnabled() const
{
    Q_D(const Workbook);
    return d->strings_to_hyperlinks_enabled;
}

void Workbook::setHtmlToRichStringEnabled(bool enable)
{
    Q_D(Workbook);
    d->html_to_richstring_enabled = enable;
}

bool Workbook::isHtmlToRichStringEnabled() const
{
    Q_D(const Workbook);
    return d->html_to_richstring_enabled;
}

QString Workbook::defaultDateFormat() const
{
    Q_D(const Workbook);
    return d->defaultDateFormat;
}

void Workbook::setDefaultDateFormat(const QString &format)
{
    Q_D(Workbook);
    d->defaultDateFormat = format;
}

/*!
 * \brief Create a defined name in the workbook.
 * \param name The defined name
 * \param formula The cell or range that the defined name refers to.
 * \param comment
 * \param scope The name of one worksheet, or empty which means golbal scope.
 * \return Return false if the name invalid.
 */

/*
bool Workbook::defineName(const QString &name, const QString &formula, const QString &comment, const QString &scope)
{
    Q_D(Workbook);

    //Remove the = sign from the formula if it exists.
    QString formulaString = formula;
    if (formulaString.startsWith(QLatin1Char('=')))
        formulaString = formula.mid(1);

    int id=-1;
    if (!scope.isEmpty()) {
        for (int i=0; i<d->sheets.size(); ++i) {
            if (d->sheets[i]->sheetName() == scope) {
                id = d->sheets[i]->sheetId();
                break;
            }
        }
    }

	d->definedNamesList.append(XlsxDefineNameData(name, formulaString, comment, id));
    return true;
}
*/

/*!
* \brief Dave added Find a named range based on the name.
* \param sName is the key value for lookup
* \return Return reference to found name range.  If not found, name is empty.
*/
DefinedName Workbook::FindNameRange(const QString &sName)
{
	return names()->find(sName);
}
//	End dave added

/*!
* \internal - Dave added
*/
/*
const QStringList Workbook::rangeNames() const
{
	Q_D(const Workbook);
	
	//	Create range name list for saves
	QStringList rangeNameList;

	QList<DefinedName*> *list = d->definedNames->getList();

	//foreach(XlsxDefineNameData data, d->definedNamesList) {
	for (qint32 i = 0; i < list->size(); ++i) 
	{
		rangeNameList.append( list->at(i)->name() );
	}

	return rangeNameList;
}
*/
/*!
* \internal
*/
AbstractSheet *Workbook::addSheet(const QString &name, AbstractSheet::SheetType type)
{
	Q_D(Workbook);
	return insertSheet(d->sheets.size(), name, type);
}

/*!
 * \internal
 */
QStringList Workbook::worksheetNames() const
{
    Q_D(const Workbook);
    return d->sheetNames;
}

/*!
 * \internal
 * Used only when load the xlsx file!!
 */
AbstractSheet *Workbook::addSheet(const QString &name, int sheetId, AbstractSheet::SheetType type)
{
    Q_D(Workbook);
    if (sheetId > d->last_sheet_id)
        d->last_sheet_id = sheetId;
    AbstractSheet *sheet=0;
    if (type == AbstractSheet::ST_WorkSheet) {
        sheet = new Worksheet(name, sheetId, this, F_LoadFromExists);
    } else if (type == AbstractSheet::ST_ChartSheet) {
        sheet = new Chartsheet(name, sheetId, this, F_LoadFromExists);
    } else {
        qWarning("unsupported sheet type.");
        Q_ASSERT(false);
    }
    d->sheets.append(QSharedPointer<AbstractSheet>(sheet));
    d->sheetNames.append(name);

    return sheet;
}

AbstractSheet *Workbook::insertSheet(int index, const QString &name, AbstractSheet::SheetType type)
{
    Q_D(Workbook);
    QString sheetName = createSafeSheetName(name);
    if(index > d->last_sheet_id){
        //User tries to insert, where no sheet has gone before.
        return 0;
    }
    if (!sheetName.isEmpty()) {
        //If user given an already in-used name, we should not continue any more!
        if (d->sheetNames.contains(sheetName))
            return 0;
    } else {
        if (type == AbstractSheet::ST_WorkSheet) {
            do {
                ++d->last_worksheet_index;
                sheetName = QStringLiteral("Sheet%1").arg(d->last_worksheet_index);
            } while (d->sheetNames.contains(sheetName));
        } else if (type == AbstractSheet::ST_ChartSheet) {
            do {
                ++d->last_chartsheet_index;
                sheetName = QStringLiteral("Chart%1").arg(d->last_chartsheet_index);
            } while (d->sheetNames.contains(sheetName));
        } else {
            qWarning("unsupported sheet type.");
            return 0;
        }
    }

    ++d->last_sheet_id;
    AbstractSheet *sheet;
    if (type == AbstractSheet::ST_WorkSheet)
        sheet = new Worksheet(sheetName, d->last_sheet_id, this, F_NewFromScratch);
    else
        sheet = new Chartsheet(sheetName, d->last_sheet_id, this, F_NewFromScratch);

    d->sheets.insert(index, QSharedPointer<AbstractSheet>(sheet));
    d->sheetNames.insert(index, sheetName);
    //d->activesheetIndex = index;
	setActiveSheetId( index );

    return sheet;
}

/*!
 * Returns current active worksheet.
 */
AbstractSheet *Workbook::activeSheet() const
{
    Q_D(const Workbook);
    if (d->sheets.isEmpty())
        const_cast<Workbook*>(this)->addSheet();
	
	//	centralized active sheet logic
	return d->sheets[activeSheetId()].data();
	//return d->sheets[d->activesheetIndex].data();
}

bool Workbook::setActiveSheet(int index)
{
    Q_D(Workbook);
    if (index < 0 || index >= d->sheets.size()) {
        //warning
        return false;
    }
    
	setActiveSheetId( index );
	//d->activesheetIndex = index;
    return true;
}

/*!
 * Rename the worksheet at the \a index to \a newName.
 */
bool Workbook::renameSheet(int index, const QString &newName)
{
    Q_D(Workbook);
    QString name = createSafeSheetName(newName);
    if (index < 0 || index >= d->sheets.size())
        return false;

    //If user given an already in-used name, return false
    for (int i=0; i<d->sheets.size(); ++i) {
        if (d->sheets[i]->sheetName() == name)
            return false;
    }

    d->sheets[index]->setSheetName(name);
    d->sheetNames[index] = name;
    return true;
}

/*!
 * Remove the worksheet at pos \a index.
 */
bool Workbook::deleteSheet(int index)
{
    Q_D(Workbook);
    if (d->sheets.size() <= 1)
        return false;
    if (index < 0 || index >= d->sheets.size())
        return false;
    d->sheets.removeAt(index);
    d->sheetNames.removeAt(index);
    return true;
}

/*!
 * Moves the worksheet form \a srcIndex to \a distIndex.
 */
bool Workbook::moveSheet(int srcIndex, int distIndex)
{
    Q_D(Workbook);
    if (srcIndex == distIndex)
        return false;

    if (srcIndex < 0 || srcIndex >= d->sheets.size())
        return false;

    QSharedPointer<AbstractSheet> sheet = d->sheets.takeAt(srcIndex);
    d->sheetNames.takeAt(srcIndex);
    if (distIndex >= 0 || distIndex <= d->sheets.size()) {
        d->sheets.insert(distIndex, sheet);
        d->sheetNames.insert(distIndex, sheet->sheetName());
    } else {
        d->sheets.append(sheet);
        d->sheetNames.append(sheet->sheetName());
    }
    return true;
}

bool Workbook::copySheet(int index, const QString &newName)
{
    Q_D(Workbook);
    if (index < 0 || index >= d->sheets.size())
        return false;

    QString worksheetName = createSafeSheetName(newName);
    if (!newName.isEmpty()) {
        //If user given an already in-used name, we should not continue any more!
        if (d->sheetNames.contains(newName))
            return false;
    } else {
        int copy_index = 1;
        do {
            ++copy_index;
            worksheetName = QStringLiteral("%1(%2)").arg(d->sheets[index]->sheetName()).arg(copy_index);
        } while (d->sheetNames.contains(worksheetName));
    }

    ++d->last_sheet_id;
    AbstractSheet *sheet = d->sheets[index]->copy(worksheetName, d->last_sheet_id);
    d->sheets.append(QSharedPointer<AbstractSheet> (sheet));
    d->sheetNames.append(sheet->sheetName());

    return false;
}

/*!
 * Returns count of worksheets.
 */
int Workbook::sheetCount() const
{
    Q_D(const Workbook);
    return d->sheets.count();
}

/*!
 * Returns the sheet object at index \a sheetIndex.
 */
AbstractSheet *Workbook::sheet(int index) const
{
    Q_D(const Workbook);
    if (index < 0 || index >= d->sheets.size())
        return 0;
    return d->sheets.at(index).data();
}

//	Dave added
Worksheet *Workbook::worksheet(int index) const
{
	Q_D(const Workbook);
	if (index < 0 || index >= d->sheets.size())
		return nullptr;

	AbstractSheet *temp = d->sheets.at(index).data();

	if (temp && temp->sheetType() == AbstractSheet::ST_WorkSheet)
		return static_cast<Worksheet *>(temp);

	//	Not found or chart.  Oops...
	return nullptr;
}

Worksheet *Workbook::worksheet(const QString& name) const
{
	//	Convert name to index and index to sheet
	return worksheet(sheetIndex(name));
}


int Workbook::sheetIndex(const QString& name) const
{
	Q_D(const Workbook);

	//	Use map for lookup by name.  Case is ignored
	//return d->sheetsMap.find(name.toUpper()).value().data();

	//	Use existing storage until testing proves it should be optimized
	return d->sheetNames.indexOf(name);
}
//	end

SharedStrings *Workbook::sharedStrings() const
{
    Q_D(const Workbook);
    return d->sharedStrings.data();
}

Styles *Workbook::styles()
{
    Q_D(Workbook);
    return d->styles.data();
}

Theme *Workbook::theme()
{
    Q_D(Workbook);
    return d->theme.data();
}

/*!
* \brief Dave added Return private data structure for direct access
* \param None
* \return Returns DefinedNames container
*/

DefinedNames* Workbook::names() const
{
	Q_D(const Workbook);
	return d->definedNames.data();
}

/*!
 * \internal
 *
 * Unlike media files, drawing file is a property of the sheet.
 */
QList<Drawing *> Workbook::drawings()
{
    Q_D(Workbook);
    QList<Drawing *> ds;
    for (int i=0; i<d->sheets.size(); ++i) {
        QSharedPointer<AbstractSheet> sheet = d->sheets[i];
        if (sheet->drawing())
        ds.append(sheet->drawing());
    }

    return ds;
}

/*!
 * \internal
 */
QList<QSharedPointer<AbstractSheet> > Workbook::getSheetsByTypes(AbstractSheet::SheetType type) const
{
    Q_D(const Workbook);
    QList<QSharedPointer<AbstractSheet> > list;
    for (int i=0; i<d->sheets.size(); ++i) {
        if (d->sheets[i]->sheetType() == type)
            list.append(d->sheets[i]);
    }
    return list;
}

void Workbook::saveToXmlFile(QIODevice *device) const
{
    Q_D(const Workbook);
    d->relationships->clear();
    if (d->sheets.isEmpty())
        const_cast<Workbook *>(this)->addSheet();

    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("workbook"));

	//	Dave-Replace head with Xl 2010+ headers
	d->saveXmlHeader( writer );

    //writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    //writer.writeAttribute(QStringLiteral("xmlns:r"), QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    writer.writeEmptyElement(QStringLiteral("fileVersion"));

	//	Dave-Replace hardcoded with original file
	QXmlStreamAttributes attrs = d->fileVersion.writeXMLAttributes();
	writer.writeAttributes( attrs );
	
	//writer.writeAttribute(QStringLiteral("appName"), QStringLiteral("xl"));
    //writer.writeAttribute(QStringLiteral("lastEdited"), QStringLiteral("4"));
    //writer.writeAttribute(QStringLiteral("lowestEdited"), QStringLiteral("4"));
    //writer.writeAttribute(QStringLiteral("rupBuild"), QStringLiteral("4505"));
//  //writer.writeAttribute(QStringLiteral("codeName"), QStringLiteral("{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}"));

    writer.writeEmptyElement(QStringLiteral("workbookPr"));

	attrs = d->workbookPr.writeXMLAttributes();
	writer.writeAttributes( attrs );

	//if (d->date1904)
    //    writer.writeAttribute(QStringLiteral("date1904"), QStringLiteral("1"));

    writer.writeStartElement(QStringLiteral("bookViews"));

	//	Dave-Use values from previous sheet
	//for( const XlsxAttributes& attrs : d->workbookView )
	//{
	//	Library assumes on single element. May need to revisit if there are multiple
	writer.writeEmptyElement( QStringLiteral( "workbookView" ) );

	attrs = d->workbookView.writeXMLAttributes();
	writer.writeAttributes( attrs );

	//}
    //writer.writeAttribute(QStringLiteral("xWindow"), QString::number(d->x_window));
    //writer.writeAttribute(QStringLiteral("yWindow"), QString::number(d->y_window));
    //writer.writeAttribute(QStringLiteral("windowWidth"), QString::number(d->window_width));
    //writer.writeAttribute(QStringLiteral("windowHeight"), QString::number(d->window_height));
    //Store the firstSheet when it isn't the default
    //For example, when "the first sheet 0 is hidden", the first sheet will be 1
    //if (d->firstsheet > 0)
    //    writer.writeAttribute(QStringLiteral("firstSheet"), QString::number(d->firstsheet + 1));
    //Store the activeTab when it isn't the first sheet
    //if (d->activesheetIndex > 0)
    //    writer.writeAttribute(QStringLiteral("activeTab"), QString::number(d->activesheetIndex));
    writer.writeEndElement();//bookViews

    writer.writeStartElement(QStringLiteral("sheets"));
    int worksheetIndex = 0;
    int chartsheetIndex = 0;
    for (int i=0; i<d->sheets.size(); ++i) {
        QSharedPointer<AbstractSheet> sheet = d->sheets[i];
        writer.writeEmptyElement(QStringLiteral("sheet"));
        writer.writeAttribute(QStringLiteral("name"), sheet->sheetName());
        writer.writeAttribute(QStringLiteral("sheetId"), QString::number(sheet->sheetId()));
        if (sheet->sheetState() == AbstractSheet::SS_Hidden)
            writer.writeAttribute(QStringLiteral("state"), QStringLiteral("hidden"));
        else if (sheet->sheetState() == AbstractSheet::SS_VeryHidden)
            writer.writeAttribute(QStringLiteral("state"), QStringLiteral("veryHidden"));

        if (sheet->sheetType() == AbstractSheet::ST_WorkSheet)
            d->relationships->addDocumentRelationship(QStringLiteral("/worksheet"), QStringLiteral("worksheets/sheet%1.xml").arg(++worksheetIndex));
        else
            d->relationships->addDocumentRelationship(QStringLiteral("/chartsheet"), QStringLiteral("chartsheets/sheet%1.xml").arg(++chartsheetIndex));

        writer.writeAttribute(QStringLiteral("r:id"), QStringLiteral("rId%1").arg(d->relationships->count()));
    }
    writer.writeEndElement();//sheets

    if (d->externalLinks.size() > 0) {
        writer.writeStartElement(QStringLiteral("externalReferences"));
        for (int i=0; i<d->externalLinks.size(); ++i) {
            writer.writeEmptyElement(QStringLiteral("externalReference"));
            d->relationships->addDocumentRelationship(QStringLiteral("/externalLink"), QStringLiteral("externalLinks/externalLink%1.xml").arg(i+1));
            writer.writeAttribute(QStringLiteral("r:id"), QStringLiteral("rId%1").arg(d->relationships->count()));
        }
        writer.writeEndElement();//externalReferences
    }

	if (!d->definedNames->empty()) {
        
		//	Moved to DefinedNames class
		d->definedNames->SaveXmlDefinedNames(writer);
    }

    writer.writeStartElement(QStringLiteral("calcPr"));
	//writer.writeAttribute(QStringLiteral("calcId"), QStringLiteral("124519"));
	
	//	Dave-Replace hardcoded with cached value from original file
	attrs = d->calcPr.writeXMLAttributes();
	writer.writeAttributes( attrs );
	
	writer.writeEndElement(); //calcPr

    writer.writeEndElement();//workbook
    writer.writeEndDocument();

    d->relationships->addDocumentRelationship(QStringLiteral("/theme"), QStringLiteral("theme/theme1.xml"));
    d->relationships->addDocumentRelationship(QStringLiteral("/styles"), QStringLiteral("styles.xml"));
    if (!sharedStrings()->isEmpty())
        d->relationships->addDocumentRelationship(QStringLiteral("/sharedStrings"), QStringLiteral("sharedStrings.xml"));
}

void WorkbookPrivate::saveXmlHeader(QXmlStreamWriter &writer) const	//	Dave added
{
	if (namespaceDeclarations.empty())
	{
		//	Sheet was created from scratch, put in basic headers
		writer.writeAttribute(QLatin1String("xmlns"), QLatin1String("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
		writer.writeAttribute(QLatin1String("xmlns:r"), QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

		//for Excel 2010 - this could be changed to workbook flag.  Assume At least Excel 2010 since 2007 is no longer supported.
		writer.writeAttribute(QLatin1String("xmlns:mc"), QLatin1String("http://schemas.openxmlformats.org/markup-compatibility/2006"));
		writer.writeAttribute(QLatin1String("mc:Ignorable"), QLatin1String("x15"));
		writer.writeAttribute(QLatin1String("xmlns:x15"), QLatin1String("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"));
	}
	else
	{
		//	If no namespace found, use xmlns
		QString tempNamespace = headerMamespace.isEmpty() ? QLatin1String("xmlns") : headerMamespace;

		//	Output from opened sheet
		QXmlStreamNamespaceDeclaration declare;
		foreach(declare, namespaceDeclarations)
		{
			//	create name with prefex
			QString prefix = tempNamespace;

			//	Skip blank, this is main namespace	
			if( !declare.prefix().isEmpty())
				prefix += QLatin1String(":") + declare.prefix().toString();

			//	Order in XL is strange, push attribute in between attributes 
			writer.writeAttribute(prefix, declare.namespaceUri().toString());

			if (prefix == QLatin1String("xmlns:mc") && headerAttributes.hasAttribute(QLatin1String("mc:Ignorable")))
			{
				//	Put ignore attribute here.  Then complete name spaces
				writer.writeAttribute(QLatin1String("mc:Ignorable"), headerAttributes.value(QLatin1String("mc:Ignorable")).toString());
			}
		}

		//	Output remaining attributes
		foreach(QXmlStreamAttribute attrib, headerAttributes)
		{
			QString name = attrib.qualifiedName().toString();

			//	Ignorable output above already
			if( name != QLatin1String("mc:Ignorable"))
				writer.writeAttribute(name, attrib.value().toString());
		}
	}
}


bool Workbook::loadFromXmlFile(QIODevice *device)
{
    Q_D(Workbook);

	//	This is inside the workbook.xml file
	QXmlStreamReader reader(device);
	while (!reader.atEnd()) {
		QXmlStreamReader::TokenType token = reader.readNext();
		if( token == QXmlStreamReader::StartElement ) {
			if( reader.name() == QLatin1String( "sheet" ) ) {
				QXmlStreamAttributes attributes = reader.attributes();
				const QString name = attributes.value( QLatin1String( "name" ) ).toString();
				int sheetId = attributes.value( QLatin1String( "sheetId" ) ).toString().toInt();
				const QString rId = attributes.value( QLatin1String( "r:id" ) ).toString();
				const QStringRef &stateString = attributes.value( QLatin1String( "state" ) );
				AbstractSheet::SheetState state = AbstractSheet::SS_Visible;
				if( stateString == QLatin1String( "hidden" ) )
					state = AbstractSheet::SS_Hidden;
				else if( stateString == QLatin1String( "veryHidden" ) )
					state = AbstractSheet::SS_VeryHidden;

				 XlsxRelationship relationship = d->relationships->getRelationshipById( rId );

				 AbstractSheet::SheetType type = AbstractSheet::ST_WorkSheet;
				 if( relationship.type.endsWith( QLatin1String( "/worksheet" ) ) )
					 type = AbstractSheet::ST_WorkSheet;
				 else if( relationship.type.endsWith( QLatin1String( "/chartsheet" ) ) )
					 type = AbstractSheet::ST_ChartSheet;
				 else if( relationship.type.endsWith( QLatin1String( "/dialogsheet" ) ) )
					 type = AbstractSheet::ST_DialogSheet;
				 else if( relationship.type.endsWith( QLatin1String( "/xlMacrosheet" ) ) )
					 type = AbstractSheet::ST_MacroSheet;
				 else
					 qWarning( "unknown sheet type" );

				 AbstractSheet *sheet = addSheet( name, sheetId, type );
				 sheet->setSheetState( state );
				 const QString fullPath = QDir::cleanPath( splitPath( filePath() )[ 0 ] + QLatin1String( "/" ) + relationship.target );
				 sheet->setFilePath( fullPath );
			 }
			 else if( reader.name() == QLatin1String( "workbookPr" ) ) {
				 d->workbookPr.readXMLAttributes( reader.attributes() );
				 //if( d->workbookPr.hasAttribute( QLatin1String( "date1904" ) ) )
				//	 d->date1904 = true;
			 }
			 /*else if( reader.name() == QLatin1String( "bookviews" ) ) {
				 while( !( reader.name() == QLatin1String( "bookviews" ) && reader.tokenType() == QXmlStreamReader::EndElement ) ) {
					 reader.readNextStartElement();
					 if( reader.tokenType() == QXmlStreamReader::StartElement ) {
						 if( reader.name() == QLatin1String( "workbookView" ) ) 
						 {
							 //	Keep list of views. Generally only one, but could be more
							 //	Library assumes a single view.  Update if actual files are different
							 //XlsxAttributes attrs;
							 //attrs.readXMLAttributes( reader.attributes());
							 d->workbookView.readXMLAttributes( reader.attributes() );
//							 if( attrs.hasAttribute( QLatin1String( "xWindow" ) ) )
//								 d->x_window = attrs.value( QLatin1String( "xWindow" ) ).toString().toInt();
//							 if( attrs.hasAttribute( QLatin1String( "yWindow" ) ) )
//								 d->y_window = attrs.value( QLatin1String( "yWindow" ) ).toString().toInt();
//							 if( attrs.hasAttribute( QLatin1String( "windowWidth" ) ) )
//								 d->window_width = attrs.value( QLatin1String( "windowWidth" ) ).toString().toInt();
//							 if( attrs.hasAttribute( QLatin1String( "windowHeight" ) ) )
//								 d->window_height = attrs.value( QLatin1String( "windowHeight" ) ).toString().toInt();
//							 if( attrs.hasAttribute( QLatin1String( "firstSheet" ) ) )
//								 d->firstsheet = attrs.value( QLatin1String( "firstSheet" ) ).toString().toInt();
//							 if( attrs.hasAttribute( QLatin1String( "activeTab" ) ) )
//								 d->activesheetIndex = attrs.value( QLatin1String( "activeTab" ) ).toString().toInt();
//						 }
//					 }
//				 }
			 }*/
			 else if( reader.name() == QLatin1String( "externalReference" ) ) {
				 QXmlStreamAttributes attributes = reader.attributes();
				 const QString rId = attributes.value( QLatin1String( "r:id" ) ).toString();
				 XlsxRelationship relationship = d->relationships->getRelationshipById( rId );

				 QSharedPointer<SimpleOOXmlFile> link( new SimpleOOXmlFile( F_LoadFromExists ) );
				 const QString fullPath = QDir::cleanPath( splitPath( filePath() )[ 0 ] + QLatin1String( "/" ) + relationship.target );
				 link->setFilePath( fullPath );
				 d->externalLinks.append( link );
			 }
			 else if( reader.name() == QLatin1String( "definedName" ) ) {
				 //	Dave-Moved to DefinedNames class
				 d->definedNames->loadXmlDefinedNames( reader );
			 }
			 else if( reader.name() == QLatin1String( "workbook" ) ) {
				 //	Dave-Moved to DefinedNames class
				 d->loadXmlHeader( reader );
			 }
			 else if( reader.name() == QLatin1String( "fileVersion" ) ) {
				 //	Dave-Cache attributes for later save
				d->fileVersion.readXMLAttributes( reader.attributes() );
			 }
			 else if( reader.name() == QLatin1String( "workbookView" ) ) {
				 //	Dave-Save complete cache of values
				d->workbookView.readXMLAttributes( reader.attributes() );
			 }
			 else if( reader.name() == QLatin1String( "calcPr" ) ) {
				//	Dave-Save complete cache of values
				d->calcPr.readXMLAttributes( reader.attributes() );
				
				//	Force calculation on open
				d->calcPr.insert( QStringLiteral( "fullCalcOnLoad" ), QStringLiteral( "1" ) );
			}
			 else if( reader.name() == QLatin1String( "extLst" ) ) {
				 //	Dave-Not interested in this value, but duplicate workbookPr element exists here.  We want to skip
				while( !( reader.name() == QLatin1String( "extLst" ) && reader.tokenType() == QXmlStreamReader::EndElement ) ) 
				{
					//	Just go through these elements			 
					 reader.readNextStartElement();
				}
			 }
		 }
    }

    return true;
}

void WorkbookPrivate::loadXmlHeader(QXmlStreamReader &reader)
{
	//	Header consists of 3 parts, name space, declaration and attributes

	//	Save name space
	headerMamespace = QLatin1String("xmlns"); // Struggling to find correct QT function

	//	Save declaration
	namespaceDeclarations = reader.namespaceDeclarations();

	//	Save header attributes
	headerAttributes = reader.attributes();
}

/*!
 * \internal
 */
QList<QSharedPointer<MediaFile> > Workbook::mediaFiles() const
{
    Q_D(const Workbook);

    return d->mediaFiles;
}

/*!
 * \internal
 */
void Workbook::addMediaFile(QSharedPointer<MediaFile> media, bool force)
{
    Q_D(Workbook);
    if (!force) {
        for (int i=0; i<d->mediaFiles.size(); ++i) {
            if (d->mediaFiles[i]->hashKey() == media->hashKey()) {
                media->setIndex(i);
                return;
            }
        }
    }
    media->setIndex(d->mediaFiles.size());
    d->mediaFiles.append(media);
}

/*!
 * \internal
 */
QList<QSharedPointer<Chart> > Workbook::chartFiles() const
{
    Q_D(const Workbook);

    return d->chartFiles;
}

/*!
 * \internal
 */
void Workbook::addChartFile(QSharedPointer<Chart> chart)
{
    Q_D(Workbook);

    if (!d->chartFiles.contains(chart))
        d->chartFiles.append(chart);
}

QT_END_NAMESPACE_XLSX
