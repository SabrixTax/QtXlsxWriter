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
#ifndef QXLSX_XLSXCELLRANGE_P_H
#define QXLSX_XLSXCELLRANGE_P_H

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
#include "xlsxcellrange.h"

#include <QSharedData>

QT_BEGIN_NAMESPACE_XLSX

class Worksheet;

class CellRangePrivate : public QSharedData
{
public:

	CellRangePrivate(Worksheet *sheet);
	
	//	Copy constructor
	CellRangePrivate(const CellRangePrivate &other);

	virtual ~CellRangePrivate();

    qint32 top, left, bottom, right;
	
	//	Range cannot exist without sheet
	Worksheet *sheet;	
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_XLSXCELLRANGE_P_H
