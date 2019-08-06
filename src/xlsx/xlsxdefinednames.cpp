/****************************************************************************
** Copyright (c) 2018 David McRaven <david.mcraven@sabrixtax.com>
** All right reserved.
**
** See Page 3956 in Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf
** See commplexType name = CT_definedName
**
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
#include "xlsxdefinednames.h"
#include "xlsxdefinednames_p.h"

//	Container for ranges
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"
#include "xlsxutility_p.h"

#include <algorithm>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

/*
#include "xlsxcell.h"
#include "xlsxcell_p.h"
#include "xlsxformat.h"
#include "xlsxformat_p.h"
#include "xlsxutility_p.h"
#include <QDateTime>
*/

QT_BEGIN_NAMESPACE_XLSX

//	Individual names
DefinedNamePrivate::DefinedNamePrivate(DefinedName *p) :
	q_ptr(p),
	category(0),
	localSheetId(0xffffffff),
	macroType(DefinedName::MacroType::XlNone),
	publishToServer(false),
	visible(true),
	workbookParameter(false),
	xlm(false),
	sheet(nullptr),
	book(nullptr),
	range(CellRange())
{
}

DefinedNamePrivate::DefinedNamePrivate(const DefinedNamePrivate * const cp) :
	name(cp->name), 
	category(cp->category),
	comment(cp->comment), 
	customMenu(cp->customMenu),
	description(cp->description),
	help(cp->help),
	localSheetId(cp->localSheetId),
	macroType(cp->macroType),
	refersTo(cp->refersTo),
	shortcutKey(cp->shortcutKey),
	statusBar(cp->statusBar),
	publishToServer(cp->publishToServer),
	visible(cp->visible),
	workbookParameter(cp->workbookParameter),
	xlm(cp->xlm),
	sheet(cp->sheet),
	book(cp->book),
	range(cp->range)
{
}

DefinedNamePrivate& DefinedNamePrivate::clone(const DefinedNamePrivate& other)
{ 
	//q_ptr=other.q_ptr;
	name				= other.name;
	category			= other.category;
	comment				= other.comment;
	customMenu			= other.customMenu;
	description			= other.description;
	help				= other.help;
	localSheetId		= other.localSheetId;
	macroType			= other.macroType;
	refersTo			= other.refersTo;
	shortcutKey			= other.shortcutKey;
	statusBar			= other.statusBar;
	publishToServer		= other.publishToServer;
	visible				= other.visible;
	workbookParameter	= other.workbookParameter;
	xlm					= other.xlm;
	sheet				= other.sheet;
	book				= other.book;
	range				= other.range;

	return *this;
}

DefinedName::DefinedName() :
	d_ptr(new DefinedNamePrivate(this))
{
}

DefinedName::DefinedName(const QString& sName, const Workbook* book) :
	d_ptr(new DefinedNamePrivate(this))
{
	Q_ASSERT(book != nullptr);
	d_ptr->name = sName;
	d_ptr->book = book;
}

/*!
* Return a copy based on this Defined name.  Performs deep copy.
*/
/*
DefinedName& DefinedName::cloneFrom(const DefinedName& other) 
{
	//Q_D(const DefinedName);
	DefinedName *aClone = new DefinedName();
	
	//	Copy data members
	aClone->d_ptr->clone(other.d_ptr);

	return *aClone;
}
*/
/*!
* \internal
*/

DefinedName::DefinedName(const DefinedName * const definedName) :
	d_ptr(new DefinedNamePrivate(definedName->d_ptr))
{
	d_ptr->q_ptr = this;
}

/*
DefinedName::DefinedName( const QString name, const QString refersTo,
	const QString comment, quint32 id, bool global, bool visible,
	DefinedName::MacroType macroType, QString shortcutKey,
	DefinedName::Category category ) :
	d_ptr(new DefinedNamePrivate(this))
{
	d_ptr->name			= name;
	d_ptr->refersTo		= refersTo;
	d_ptr->comment		= comment;
	d_ptr->sheetId		= id;
	d_ptr->isGlobal		= global;
	d_ptr->isVisible	= visible;
	d_ptr->macroType	= macroType;
	d_ptr->shortcutKey	= shortcutKey;
	d_ptr->category		= category;
}
*/

/*!
* Destroys the Cell and cleans up.
*/
DefinedName::~DefinedName()
{
	delete d_ptr;
}


QString DefinedName::name() const
{
	Q_D(const DefinedName);
	return d->name;
}

void DefinedName::setName(const QString& value)
{
	Q_D(DefinedName);
	d->name = value;
}

quint16 DefinedName::category() const
{
	Q_D(const DefinedName);
	return d->category;
}

void DefinedName::setCategory(quint16 value)
{
	Q_D(DefinedName);
	d->category = value;
}

QString DefinedName::comment() const
{
	Q_D(const DefinedName);
	return d->comment;
}

void DefinedName::setComment(const QString& value)
{
	Q_D(DefinedName);
	d->comment = value;
}

QString DefinedName::customMenu() const
{
	Q_D(const DefinedName);
	return d->customMenu;
}

void DefinedName::setCustomMenu(const QString& value)
{
	Q_D(DefinedName);
	d->customMenu = value;
}

QString DefinedName::description() const
{
	Q_D(const DefinedName);
	return d->description;
}

void DefinedName::setDescription(const QString& value)
{
	Q_D(DefinedName);
	d->description = value;
}

QString DefinedName::help() const
{
	Q_D(const DefinedName);
	return d->help;
}

void DefinedName::setHelp(const QString& value)
{
	Q_D(DefinedName);
	d->help = value;
}

bool DefinedName::isGlobal() const
{
	Q_D(const DefinedName);
	return d->localSheetId == 0xffffffff;
}

quint32 DefinedName::localSheetId() const
{
	Q_D(const DefinedName);
	return d->localSheetId;//	d->sheet == nullptr ? 0xffffffff : d->sheet->sheetId();
}

void DefinedName::setLocalSheetId(quint32 value)
{
	Q_D(DefinedName);
	d->localSheetId = value;
}

DefinedName::MacroType DefinedName::macroType() const
{
	Q_D(const DefinedName);
	return d->macroType;
}

void DefinedName::setMacroType(MacroType value)
{
	Q_D(DefinedName);
	d->macroType = value;
}

QString DefinedName::refersTo() const
{
	Q_D(const DefinedName);
	return d->refersTo;
}

void DefinedName::setRefersTo(const QString& value)
{
	Q_D(DefinedName);
	d->refersTo = value;
	QString temp = value;

	//	Translate to pointer for range functions
	//	Split ranges out
	QStringList ranges = value.split(QLatin1Char(','));
		
	//	Deal with only first range for now
	QStringList parts = ranges[0].split(QLatin1Char('!'));
		
	//	Get last part. If no sheet name, this will be only part
	temp = parts.at(parts.count() - 1);

	//	If missing sheet, get from RefersTo
	if (!d->sheet)
	{
		QString smartName = parts[0];

		//	See if there is valid pointer
		if (smartName.startsWith(QLatin1Char('\'')) && smartName.endsWith(QLatin1Char('\'')))
		{
			//	Eliminate quotes for lookup
			smartName = unescapeSheetName(smartName);
		}

		d->sheet = d->book->worksheet(smartName);
	}

	//d->range = CellRange(temp, d->sheet);
	d->range = CellRange(temp);
	d->range.setSheet(d->sheet);
}

bool DefinedName::publishToServer() const
{
	Q_D(const DefinedName);
	return d->publishToServer;
}

void DefinedName::setPublishToServer(bool value)
{
	Q_D(DefinedName);
	d->publishToServer = value;
}

QString DefinedName::shortcutKey() const
{
	Q_D(const DefinedName);
	return d->shortcutKey;
}

void DefinedName::setShortcutKey(const QString& value)
{
	Q_D(DefinedName);
	//	This should be single character
	assert(value.length() < 2);
	d->shortcutKey = value;
}

QString DefinedName::statusBar() const
{
	Q_D(const DefinedName);
	return d->statusBar;
}

void DefinedName::setStatusBar(const QString& value)
{
	Q_D(DefinedName);
	d->statusBar = value;
}

bool DefinedName::visible() const
{
	Q_D(const DefinedName);
	return d->visible;
}

void DefinedName::setVisible(bool value)
{
	Q_D(DefinedName);
	d->visible = value;
}

bool DefinedName::workbookParameter() const
{
	Q_D(const DefinedName);
	return d->workbookParameter;
}

void DefinedName::setWorkbookParameter(bool value)
{
	Q_D(DefinedName);
	d->workbookParameter = value;
}

bool DefinedName::xlm() const
{
	Q_D(const DefinedName);
	return d->xlm;
}

void DefinedName::setXlm(bool value)
{
	Q_D(DefinedName);
	d->xlm = value;
}

const Worksheet* DefinedName::sheet() const
{
	Q_D(const DefinedName);
	return d->sheet;
}

void DefinedName::setSheet(Worksheet* sheet)
{
	Q_D(DefinedName);
	d->sheet = sheet;
}

const Workbook* DefinedName::book() const
{
	Q_D(const DefinedName);
	return d->book;
}

/*
void DefinedName::setBook( Workbook *book)
{
	Q_D(DefinedName);
	d->book = book;
}
*/
const CellRange DefinedName::range() const
{
	Q_D(const DefinedName);
	return d->range;
}

void DefinedName::setRange(const CellRange& range)
{
	Q_D(DefinedName);
	d->range = range;
}

//	Names container
DefinedNamesPrivate::DefinedNamesPrivate(DefinedNames *p, Workbook *workbook) :
    q_ptr(p), book(workbook)
{
}

DefinedNamesPrivate::DefinedNamesPrivate(const DefinedNamesPrivate * const /*cp*/)
    //: value(cp->value), formula(cp->formula), cellType(cp->cellType)
    //, format(cp->format), richString(cp->richString), parent(cp->parent)
{

}

DefinedNamesPrivate::~DefinedNamesPrivate()
{
	qDeleteAll(nameMap);
}

/*!
 * \internal
 * Created by Worksheet only.
 */
DefinedNames::DefinedNames(Workbook *workbook) :
    d_ptr(new DefinedNamesPrivate(this, workbook))
{
}

/*!
 * \internal
 */
DefinedNames::DefinedNames(const DefinedNames * const name):
    d_ptr(new DefinedNamesPrivate(name->d_ptr))
{
    d_ptr->q_ptr = this;
}

/*!
 * Destroys the Cell and cleans up.
 */
DefinedNames::~DefinedNames()
{
    delete d_ptr;
}

/*!
 * Return state of container
 */

bool DefinedNames::empty() const
{
    Q_D(const DefinedNames);
    return d->nameMap.isEmpty();
}

bool DefinedNames::append(DefinedName* namedRange) 
{
	Q_D(DefinedNames);
	
	//	Make copy for container, strip off xlnm for internal ranges
	QString key = namedRange->name().remove("_xlnm.").toUpper();

	//	Check for local ranges, they have special naming rules
	if (!namedRange->isGlobal())
	{
		//	Append sheet ID for local ranges
		key += QString("-%1").arg(namedRange->localSheetId());
	}

	//	Add new range, search by name
	QMap<QString, DefinedName*>::ConstIterator it = 
		d->nameMap.insert(key, namedRange);

	//	Clear cached result
	if (!d->cachedList.empty())
	{
		d->cachedList.clear();
	}

	return it != d->nameMap.constEnd();
}

/*
bool DefinedNames::append(const QString& name, const QString& formula, const QString& comment, const QString &scope)
{
	//	Note that global value has no meaning without a sheet ID.  Ignored for now
	Q_D(DefinedNames);
	
	DefinedName *namedRange = new DefinedName(name, d->book);
	namedRange->setRefersTo(formula);
	namedRange->setComment(comment);

	//	Each DefinedName has pointer to workbook
	//namedRange->setBook(d->book);

	//	A range may have multiple selections on different tabs
	//	But a range can only be local scope on 1 sheet
	const int index = d->book->sheetIndex(scope);
	
	//	If valid tab was found, save 
	if( index > 0)
	{ 

		Worksheet *sheet = static_cast<Worksheet*>(d->book->sheet(index));
		Q_ASSERT(sheet != nullptr);
		namedRange->setSheet(sheet);
	}


	//namedRange->setSheet(sheet);

	QMap<QString, DefinedName*>::const_iterator it =
		d->nameMap.insert(namedRange->name(), namedRange);

	//	Clear cached result
	if (!d->cachedList.empty())
	{
		d->cachedList.clear();
	}

	return it != d->nameMap.constEnd();
}
*/

DefinedName* DefinedNames::find(const QString& key) const
{
	Q_D(const DefinedNames);

	QMap<QString, DefinedName*>::const_iterator it = d->nameMap.find(key.toUpper());
	if (it != d->nameMap.end()) return it.value();

	//	Not found
	return nullptr;
}

QList<DefinedName*> DefinedNames::getList() const
{
	Q_D(const DefinedNames);

	return d->nameMap.values();
}

const QStringList& DefinedNames::getDefinedNameList() const
{
	Q_D(const DefinedNames);

	//	If cache is cleared, rebuild
	if (d->cachedList.empty())
	{
		QMap<QString, DefinedName*>::ConstIterator it = d->nameMap.constBegin();

		while (it != d->nameMap.constEnd())
		{
			const DefinedName* name = it.value();

			//	Skip if missing sheet, this is a constant.  Skip if hidden or invalid range
			if (name->sheet() && name->visible() && name->range().isValid())
			{
				QString rangeName;
				if (!name->isGlobal())
				{
					//	Local ranges are included in range name
					rangeName = name->sheet()->sheetName();

					//	Look for embedded space.  This is quick fix
					if (name->sheet()->sheetName().contains(" "))
					{
						//	Escape sheet name with space with ' '
						rangeName = escapeSheetName(name->sheet()->sheetName()); 
					}

					//	Sheet is separated by ! from address
					rangeName += +"!";
				}

				//	Internal ranges are prepended with hidden name.  Remove it
				rangeName += it.value()->name().remove("_xlnm.");

				//	Copy values from map.  Already in sorted alpha order
				d->cachedList.append(rangeName);
			}
			++it;
		}
	}

	//	Names cannot exist without workbook.  
	return d->cachedList;
}

quint32 DefinedNames::count() const
{
	Q_D(const DefinedNames);
	return d->nameMap.count();
}

void DefinedNames::SaveXmlDefinedNames(QXmlStreamWriter &writer)
{
	Q_D(DefinedNames);

	writer.writeStartElement(QStringLiteral("definedNames"));

	//	Get list and sort by name
	QList<DefinedName*> list = getList();

	//foreach(XlsxDefineNameData data, d->definedNamesList) {
	for (qint32 i = 0; i < list.size(); ++i) 
	{
		const DefinedName* pData = list.at(i);
		writer.writeStartElement(QStringLiteral("definedName"));
		writer.writeAttribute(QStringLiteral("name"), pData->name());
		if (!pData->comment().isEmpty())
			writer.writeAttribute(QStringLiteral("comment"), pData->comment());
	
		//	Only saved for xlsm file
		if (!pData->customMenu().isEmpty())
			writer.writeAttribute(QStringLiteral("customMenu"), pData->customMenu());
		if (!pData->description().isEmpty())
			writer.writeAttribute(QStringLiteral("description"), pData->description());
		if (!pData->help().isEmpty())
			writer.writeAttribute(QStringLiteral("help"), pData->help());
		if (!pData->shortcutKey().isEmpty())
			writer.writeAttribute(QStringLiteral("shortcutKey"), pData->shortcutKey());
		if (!pData->statusBar().isEmpty())
			writer.writeAttribute(QStringLiteral("statusBar"), pData->statusBar());
		//	End xlsm save
	
		//	Convert from VBA syntax to XML file format
		if (pData->macroType() != DefinedName::MacroType::XlNone)
		{
			//	Can only be one or the other
			if (pData->macroType() == DefinedName::MacroType::XlCommand)
			{
				writer.writeAttribute(QStringLiteral("vbProcedure"), "1");
			}
			if (pData->macroType() == DefinedName::MacroType::XlFunction)
			{
				writer.writeAttribute(QStringLiteral("function"), "1");
			}
	
			writer.writeAttribute(QStringLiteral("xlm"), "1");
		}
		else
		{
			//	Corner case where function/command isn't set, but range is treated as macro
			if (pData->xlm())
			writer.writeAttribute(QStringLiteral("xlm"), "1");
		}

		if (pData->category() > 0)
			writer.writeAttribute(QStringLiteral("functionGroupId"), QString::number(pData->category()));
		if (pData->publishToServer())
			writer.writeAttribute(QStringLiteral("publishToServer"), "1");

		//	VBA uses visible attribute. XML uses hidden flag.  Convert
		if (!pData->visible())
			writer.writeAttribute(QStringLiteral("hidden"), "1");
		if (pData->workbookParameter())
			writer.writeAttribute(QStringLiteral("workbookParameter"), "1");


		if (pData->localSheetId() > 0) 
		{
			//find the local index of the sheet.
			for (int i=0; i<d->book->sheetCount(); ++i) {
				if (d->book->sheet(i)->sheetId() == pData->localSheetId()) 
				{
					writer.writeAttribute(QStringLiteral("localSheetId"), QString::number(i));
					break;
				}
			}
		}
		writer.writeCharacters(pData->refersTo());
		writer.writeEndElement();//definedName
	}
	writer.writeEndElement();//definedNames
}

void DefinedNames::loadXmlDefinedNames(QXmlStreamReader &reader)
{
	Q_D(DefinedNames);

	Q_ASSERT(reader.name() == QLatin1String("definedName"));
	
	// Dave See Page 3956 in Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf
	// See commplexType name = CT_definedName
	QXmlStreamAttributes attrs = reader.attributes();
	QString name = attrs.value(QLatin1String("name")).toString();
	DefinedName *data = new DefinedName(name, d->book);

	if (attrs.hasAttribute(QLatin1String("comment")))
		data->setComment(attrs.value(QLatin1String("comment")).toString());
	if (attrs.hasAttribute(QLatin1String("localSheetId"))) {
		quint32 localId = attrs.value(QLatin1String("localSheetId")).toString().toInt();

		//	Assume local sheet is same as worksheet
		Worksheet* tempSheet = d->book->worksheet(localId);
		data->setSheet(tempSheet);
		data->setLocalSheetId(tempSheet->sheetId());
	}

	//	Added additional potental fields
	//	Called functionGroupID in file.  VBA calls this Category.
	if (attrs.hasAttribute(QLatin1String("functionGroupId")))
		data->setCategory(attrs.value(QLatin1String("functionGroupId")).toString().toInt());

	//	Straight copy-These are xlsm specific feilds
	if (attrs.hasAttribute(QLatin1String("customMenu")))
		data->setCustomMenu(attrs.value(QLatin1String("customMenu")).toString());
	if (attrs.hasAttribute(QLatin1String("description")))
		data->setDescription(attrs.value(QLatin1String("description")).toString());
	if (attrs.hasAttribute(QLatin1String("help")))
		data->setHelp(attrs.value(QLatin1String("help")).toString());
	if (attrs.hasAttribute(QLatin1String("publishToServer")))
		data->setPublishToServer(attrs.value(QLatin1String("publishToServer")).toInt() == 1);
	if (attrs.hasAttribute(QLatin1String("shortcutKey")))
		data->setShortcutKey(attrs.value(QLatin1String("shortcutKey")).toString());
	if (attrs.hasAttribute(QLatin1String("statusBar")))
		data->setStatusBar(attrs.value(QLatin1String("statusBar")).toString());

	//	Database saves as hidden flag. VBA uses isVisible. Flip bool flag.
	if (attrs.hasAttribute(QLatin1String("hidden")))
		data->setVisible(attrs.value(QLatin1String("hidden")).toInt() != 1);

	//	Saved as series of bools in file.  VBA uses enum.  Remap
	if (attrs.hasAttribute(QLatin1String("function")))
	{
		if (attrs.value(QLatin1String("function")).toInt() == 1)
			data->setMacroType(DefinedName::MacroType::XlFunction);
	}

	if (attrs.hasAttribute(QLatin1String("vbProcedure")))
	{
		if (attrs.value(QLatin1String("vbProcedure")).toInt() == 1)
			data->setMacroType(DefinedName::MacroType::XlCommand);
	}

	//	This field is automatically true if attribute function or vbProcedure is true. No need to capture
	//	Some excel files have this attribute set even though not command or function.  Keep for file consistency.
	if (attrs.hasAttribute(QLatin1String("xlm")))
		data->setXlm(attrs.value(QLatin1String("xlm")).toInt() == 1);
	if (attrs.hasAttribute(QLatin1String("workbookParameter")))
		data->setWorkbookParameter(attrs.value(QLatin1String("workbookParameter")).toInt() == 1);

	data->setRefersTo(reader.readElementText());

	//	Store, don't copy
	append(data);
	
}

QT_END_NAMESPACE_XLSX
