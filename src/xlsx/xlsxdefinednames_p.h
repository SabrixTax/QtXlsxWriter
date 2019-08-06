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
#ifndef QXLSX_XLSDEFINENAME_P_H
#define QXLSX_XLSDEFINENAME_P_H

//
//  W A R N I N G
//  -------------
//
// This file is not part of the Qt Xlsx API.  It exists for the convenience
// of the Qt Xlsx.  This header file may change from
// version to version without notice, or even be removed.
//
// We mean it.
//

#include "xlsxglobal.h"
#include "xlsxdefinednames.h"
#include "xlsxcellrange.h"

#include <QString>
#include <QMap>
#include <QList>
#include <QSharedPointer>

QT_BEGIN_NAMESPACE_XLSX

class Workbook;
/*
struct XlsxDefineNameData
{
	XlsxDefineNameData()
		:sheetId(-1)
	{}
	XlsxDefineNameData(const QString &name, const QString &formula, const QString &comment, int sheetId = -1)
		:name(name), formula(formula), comment(comment), sheetId(sheetId), bGlobalScope(false)
	{

	}
	bool bGlobalScope;
	QString name;
	QString formula;
	QString comment;
	//using internal sheetId, instead of the localSheetId(order in the workbook)
	quint32 sheetId;	//	We always want ID where formula is located
};
*/

class DefinedNamePrivate
{
	Q_DECLARE_PUBLIC(DefinedName)
public:
	DefinedNamePrivate(DefinedName *p);
	DefinedNamePrivate(const DefinedNamePrivate * const cp);
	DefinedNamePrivate& clone(const DefinedNamePrivate& other);

	bool						publishToServer;
	bool						visible;
	bool						workbookParameter;
	bool						xlm;

	QString						name;
	QString						comment;

	//	xlsm only fields
	QString						customMenu;
	QString						description;
	QString						help;
	QString						refersTo;	//	Changed name to match MS documentation
	QString						shortcutKey;
	QString						statusBar;

	//using internal sheetId, instead of the localSheetId(order in the workbook)
	quint16						category;
	quint32						localSheetId;	//	We always want ID where formula is located
	DefinedName::MacroType		macroType;
	//	Cached state-keep pointers to data to avoid look-ups on each use

	Worksheet *sheet;
	const Workbook *book;
	DefinedName *q_ptr;
	
	//	Initial implementation only addresses single range
	//	Defined Name can have multiple ranges
	CellRange range;
private:

};

class DefinedNamesPrivate
{
    Q_DECLARE_PUBLIC(DefinedNames)
public:
	DefinedNamesPrivate(DefinedNames *p, Workbook *workbook);
	DefinedNamesPrivate(const DefinedNamesPrivate * const cp);
	~DefinedNamesPrivate();

	QMap<QString, DefinedName*>	nameMap;
	//mutable QList<DefinedName*>	cachedList;	//	This list does not change internal state, only for lookup
	mutable QStringList			cachedList;	//	This list does not change internal state, only for lookup
	DefinedNames				*q_ptr;
	const Workbook				*book;

private:

};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_XLSDEFINENAME_P_H
