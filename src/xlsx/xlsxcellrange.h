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
#ifndef QXLSX_XLSXCELLRANGE_H
#define QXLSX_XLSXCELLRANGE_H

#include "xlsxglobal.h"

#include "xlsxcellreference.h"
#include "xlsxcellrange_p.h"	//	Not happy about this
#include "xlsxformat.h"

#include <QSharedData>
#include <QVariant>
#include <QString>
//#include <QExplicitlySharedDataPointer>

QT_BEGIN_NAMESPACE_XLSX

class CellRangePrivate;
class Worksheet;
class Cell;
class Format;

class Q_XLSX_EXPORT CellRange
{
public:
	enum RangeSearch : quint8
	{
		xlDown = 0X1,
		xlUp = 0X2,
		xlColumn = xlDown | xlUp,
		xlLeft = 0x04,
		xlLeftDown = xlLeft | xlDown,
		xlLeftUp = xlLeft | xlUp,
		xlRight = 0x08,
		xlRightDown = xlRight | xlDown,
		xlRightUp = xlRight | xlUp,
		xlRow = xlLeft | xlRight,
		xlDockTop = xlRow | xlUp,
		xlDockBottom = xlRow | xlDown,
		xlDocLeft = xlColumn | xlLeft,
		xlDocRight = xlColumn | xlRight,
		xlAll = xlRow | xlColumn
	};

	CellRange(Worksheet *sheet = nullptr);
	CellRange(int firstRow, int firstColumn, int lastRow, int lastColumn, Worksheet *sheet = nullptr);
    CellRange(const CellReference &topLeft, const CellReference &bottomRight, Worksheet *sheet = nullptr);
    CellRange(const QString &range, Worksheet *sheet = nullptr);
    CellRange(const char *range, Worksheet* sheet);
    
	//	Copy constructor-COW here
	CellRange(const CellRange &other);
    ~CellRange();

    QString toString(bool row_abs=false, bool col_abs=false) const;
    bool isValid() const;

	//	Getter/setters
	void setFirstRow(int row); 
	void setLastRow(int row);
	void setFirstColumn(int col);
	void setLastColumn(int col); 
	qint32 firstRow() const; 
	qint32 lastRow() const; 
	qint32 firstColumn() const; 
	qint32 lastColumn() const; 
	qint32 rowCount() const;
	qint32 columnCount() const;
	CellReference topLeft() const; 
	CellReference topRight() const; 
	CellReference bottomLeft() const; 
	CellReference bottomRight() const;

	//	Dave additions
	//	set spreadsheet
	void setSheet(Worksheet *sheet);

	//	Return offsets copy of range with given offset
	CellRange	offset( int row,  int col) const;
	
	//	Return range from this position to end in selected direction(s)
	CellRange	end(RangeSearch direction = xlDown) const;
	
	//	Refer to actual cell in sheet
	Cell*		cellAt(int row, int col) const;

	//	Read data
	QVariant read(int row, int col) const;

	//	Use Excel type logic to return text value
	QString value2(int row, int col) const;

	//	Write operations
	bool write(int row, int col, const QString& value, const Format& format=Format() );
	bool writeString(int row, int column, const QString &value, const Format &format=Format());
	Format columnFormat( int col ) const;
	bool clear(int row, int col);

	bool operator==(const CellRange& rhs) const;
	bool operator!=(const CellRange& rhs) const;

private:
	//	Private initializatin
	void init(const QString &range);

	QSharedDataPointer<CellRangePrivate> d;
};

//	Not visible outside dll
//bool operator==(const CellRange& lhs, const CellRange& rhs);
//bool operator!=(const CellRange& lhs, const CellRange& rhs);

QT_END_NAMESPACE_XLSX

Q_DECLARE_TYPEINFO(QXlsx::CellRange, Q_MOVABLE_TYPE);

#endif // QXLSX_XLSXCELLRANGE_H
