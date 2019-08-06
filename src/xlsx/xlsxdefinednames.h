/****************************************************************************
** Copyright (c) 2018 David McRaven <david.mcraven@sabrixtax.com>
** All right reserved.
**
** See Page 3956 in Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf
** See commplexType name = CT_definedName
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
#ifndef QXLSX_XLSDEFINENAME_H
#define QXLSX_XLSDEFINENAME_H

#include "xlsxglobal.h"
#include <QChar>
//#include "xlsxformat.h"
//#include <QVariant>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>

QT_BEGIN_NAMESPACE_XLSX

class DefinedNamePrivate;
class DefinedNamesPrivate;
class Workbook;
class Worksheet;
class CellRange;
class AbstractSheet;
//class QXmlStreamReader;	- Get redefinition error, must use include
//class QXmlStreamWriter;	- Get redefinition error, must use include

//class Format;
//class CellFormula;
//class CellPrivate;
//class WorksheetPrivate;

class Q_XLSX_EXPORT DefinedName
{
	Q_DECLARE_PRIVATE(DefinedName)
public:
	enum MacroType : quint8
	{
		XlNone=0,
		XlCommand,
		XlFunction
	};

	enum Category : quint16
	{
		All=0,
		DateAndTime,
		MathAndTrig,
		Statistical,
		LookupAndReference,
		Database,
		Text,
		Logical,
		information,
		Commands,
		Customizing,
		MacroControl,
		External,
		UserDefined,
		Engineering,
		Cube
	};

	DefinedName(const DefinedName * const definedName);
	//const QString refersTo,
	//const QString comment, bool global); 
	//	quint32 id, QString shortcutKey,
//		bool visible = true,
//		DefinedName::MacroType macroType = DefinedName::MacroType::XlNone,
//		quint16 category = 0);
	~DefinedName();

	QString name() const;
	void setName(const QString& value);
	quint16 category() const;
	void setCategory(quint16 value);
	QString comment() const;
	void setComment(const QString& value);
	QString customMenu() const;
	void setCustomMenu(const QString& value);
	QString description() const;
	void setDescription(const QString& value);
	QString help() const;
	void setHelp(const QString& value);
	quint32 localSheetId() const;
	MacroType macroType() const;
	void setMacroType(MacroType value);
	bool publishToServer() const;
	void setPublishToServer(bool value);
	QString shortcutKey() const;
	void setShortcutKey(const QString& value);
	QString statusBar() const;
	void setStatusBar(const QString& value);
	bool visible() const;
	void setVisible(bool value);
	bool workbookParameter() const;
	void setWorkbookParameter(bool value);
	bool xlm() const;
	void setXlm(bool value);

	//	Formula or constant
	QString refersTo() const;
	void setRefersTo(const QString& value);

	//	Link to owned sheet. Null if workbook
	const Worksheet* sheet() const;
	void setSheet(Worksheet* sheet);
	
	bool isGlobal() const;

	//	Link to owned sheet. Null if workbook
	const Workbook* book() const;

	//	If not constant, this will have a sheet range
	const CellRange range() const;
	void setRange(const CellRange& range);

	//	Helper function to create copy
	//static DefinedName& cloneFrom(const DefinedName& other);

private:
	Q_DISABLE_COPY(DefinedName);

	//	Needed so that defined names can create with pointer to workbook
	friend class DefinedNames;

	//	DefinedNames creates names
	DefinedName();
	DefinedName(const QString& name, const Workbook* book);

	//	Defined Names creates each DefineName
	//void setBook(Workbook* book);
	void setLocalSheetId(quint32 value);

	DefinedNamePrivate * const d_ptr;
};

class Q_XLSX_EXPORT DefinedNames
{
    Q_DECLARE_PRIVATE(DefinedNames)
public:
	//	Default constructor
	DefinedNames(Workbook *workbook);
	~DefinedNames();

	bool empty() const;
	bool append(DefinedName* namedRange);
	DefinedName* find(const QString& key) const;
	quint32 count() const;
	QList<DefinedName*> getList() const;
	const QStringList& getDefinedNameList() const;

	bool append(const QString& name, const QString& formula, const QString& comment, const QString &scope);

	//	Persistence functions
	void SaveXmlDefinedNames(QXmlStreamWriter &writer);
	void loadXmlDefinedNames(QXmlStreamReader &reader);

private:
	Q_DISABLE_COPY(DefinedNames);

	DefinedNames(const DefinedNames * const definedNames );
	DefinedNamesPrivate * const d_ptr;
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_XLSDEFINENAME_H
