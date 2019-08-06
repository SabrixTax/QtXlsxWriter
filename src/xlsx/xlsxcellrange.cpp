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
#include "xlsxcellrange.h"
#include "xlsxcellrange_p.h"	//	Dave added-implement COW
#include "xlsxcellreference.h"
#include "xlsxcell.h"			//	Dave added-Return cell location
#include "xlsxworksheet.h"		//	Dave added-Pointer to parent
#include "xlsxworksheet_p.h"	//	Dave-For max sheet dimensions
#include <QString>
#include <QPoint>
#include <QStringList>

QT_BEGIN_NAMESPACE_XLSX

CellRangePrivate::CellRangePrivate(Worksheet *sheet) :
	sheet(sheet), top(-1), left(-1), bottom(-2), right(-2)
{

}

CellRangePrivate::CellRangePrivate(const CellRangePrivate &other) :
	QSharedData(other), sheet(other.sheet), top(other.top),
	left(other.left), bottom(other.bottom), right(other.right) 
{
}


CellRangePrivate::~CellRangePrivate() 
{
}

/*	caused link error with other exe/dll
bool operator ==(const CellRangePrivate &lhs, const CellRangePrivate &rhs)
{
	//	If same address, data is same
	if (&lhs == &rhs) return true;

	//	Punt to private data compare
	return	lhs.sheet == rhs.sheet &&
			lhs.top == rhs.top &&
			lhs.bottom == rhs.bottom &&
			lhs.left == rhs.left &&
			lhs.right == rhs.right;
}

bool operator !=(const CellRangePrivate &lhs, const CellRangePrivate &rhs)
{
	return !(lhs == rhs);
}
*/
/*!
    \class CellRange
    \brief For a range "A1:B2" or single cell "A1"
    \inmodule QtXlsx

    The CellRange class stores the top left and bottom
    right rows and columns of a range in a worksheet.
*/

/*!
    Constructs an range, i.e. a range
    whose rowCount() and columnCount() are 0.
*/

CellRange::CellRange(Worksheet *sheet)
	:d(new CellRangePrivate(sheet))
{
}

/*!
    Constructs the range from the given \a top, \a
    left, \a bottom and \a right rows and columns.

    \sa topRow(), leftColumn(), bottomRow(), rightColumn()
*/
CellRange::CellRange(int top, int left, int bottom, int right, Worksheet* sheet)
	:d(new CellRangePrivate(sheet))
{
	//	Set data members
	d->top		= top;
	d->left		= left;
	d->bottom	= bottom;
	d->right	= right;
}


CellRange::CellRange(const CellReference &topLeft, const CellReference &bottomRight, Worksheet *sheet)
	:d(new CellRangePrivate(sheet))
{
	//	Set data members
	d->top		= topLeft.row();
	d->left		= topLeft.column();
	d->bottom	= bottomRight.row();
	d->right	= bottomRight.column();
}

/*!
    \overload
    Constructs the range form the given \a range string.
*/
CellRange::CellRange(const QString &range, Worksheet* sheet) 
	:d(new CellRangePrivate(sheet))
{
	init(range);
}

/*!
    \overload
    Constructs the range form the given \a range string.
*/
CellRange::CellRange(const char *range, Worksheet* sheet)
	:d(new CellRangePrivate(sheet))
{
	init(QString::fromLatin1(range));
}

void CellRange::init(const QString &range)
{
	QString temp = range;

	//	Dave-Add logic to pull out sheet name
	if (range.contains(QLatin1Char('!')))
	{
		//	We need to split out sheet data
		QStringList rs = range.split(QLatin1Char('!'));
		Q_ASSERT(rs.size() < 3);
		
		//	Parse remainder
		temp = rs[1];
	}

    QStringList rs = temp.split(QLatin1Char(':'));
    if (rs.size() == 2) {
        CellReference start(rs[0]);
        CellReference end(rs[1]);
		d->top = start.row();
		d->left = start.column();
		d->bottom = end.row();
		d->right = end.column();
    } else {
        CellReference p(rs[0]);
		d->top = p.row();
		d->left = p.column();
		d->bottom = p.row();
		d->right = p.column();
    }
}

/*!
    Constructs a the range by copying the given \a
    other range.
*/
CellRange::CellRange(const CellRange &other)
    : d(other.d)
{
}

/*!
    Destroys the range.
*/
CellRange::~CellRange()
{
}

/*!
     Convert the range to string notation, such as "A1:B5".
*/
QString CellRange::toString(bool row_abs, bool col_abs) const
{
    if (!isValid())
        return QString();

    if (d->left == d->right && d->top == d->bottom) {
        //Single cell
        return CellReference(d->top, d->left).toString(row_abs, col_abs);
    }

    QString cell_1 = CellReference(d->top, d->left).toString(row_abs, col_abs);
    QString cell_2 = CellReference(d->bottom, d->right).toString(row_abs, col_abs);
    return cell_1 + QLatin1String(":") + cell_2;
}

//	CellRange getter/setter functions
void CellRange::setFirstRow(int row) { d->top = row; }
void CellRange::setLastRow(int row) { d->bottom = row; }
void CellRange::setFirstColumn(int col) { d->left = col; }
void CellRange::setLastColumn(int col) { d->right = col; }
qint32 CellRange::firstRow() const { return d->top; }
qint32 CellRange::lastRow() const { return d->bottom; }
qint32 CellRange::firstColumn() const { return d->left; }
qint32 CellRange::lastColumn() const { return d->right; }
qint32 CellRange::rowCount() const { return d->bottom - d->top + 1; }
qint32 CellRange::columnCount() const { return d->right - d->left + 1; }
CellReference CellRange::topLeft() const { return CellReference(d->top, d->left); }
CellReference CellRange::topRight() const { return CellReference(d->top, d->right); }
CellReference CellRange::bottomLeft() const { return CellReference(d->bottom, d->left); }
CellReference CellRange::bottomRight() const { return CellReference(d->bottom, d->right); }

bool CellRange::operator==(const CellRange &rhs) const
{
	//	Test if data is point to same address, if so, this is shared and equal
	if (&this->d == &rhs.d) return true;

	//	compare data members
	return	this->d->sheet == rhs.d->sheet &&
			this->d->top == rhs.d->top &&
			this->d->bottom == rhs.d->bottom &&
			this->d->left == rhs.d->left &&
			this->d->right == rhs.d->right;
}
bool CellRange::operator!=(const CellRange &rhs) const
{
	//	Dave-just reimplement as opposite of equality
	return !(*this == rhs);
}

/*
bool operator ==(const CellRange &lhs, const CellRange &rhs) 
{
	//	Test if data is point to same address, if so, this is shared and equal
	if (&lhs.d == &rhs.d) return true;

	//	compare data members
	return	lhs.d->sheet	== rhs.d->sheet &&
			lhs.d->top		== rhs.d->top &&
			lhs.d->bottom	== rhs.d->bottom &&
			lhs.d->left		== rhs.d->left &&
			lhs.d->right	== rhs.d->right;
}
bool operator !=(const CellRange &lhs, const CellRange &rhs) 
{
	//	Dave-just reimplement as opposite of equality
	return !(lhs == rhs);
}
*/

/*!
 * Returns true if the Range is valid.
 */
bool CellRange::isValid() const
{
    return d->left <= d->right && d->top <= d->bottom;
}

CellRange CellRange::offset( int row,  int col) const
{
	CellRange temp(*this);

	if (row != 0)
	{
		//	protect against large negative offsets
		temp.d->top += row;
		temp.d->bottom += row;
	}

	if (col != 0)
	{
		temp.d->left += col;
		temp.d->right += col;
	}
	
	//	Use COW object
	return temp;
}

QVariant CellRange::read(int row, int col) const
{
	//	Delegate to sheet if present.  
	//	Offsets are relative to upper left cell range
	if(!d->sheet) return  QVariant();
	
	return d->sheet->read(d->top + row, d->left + col);
}

bool CellRange::write(int row, int col, const QString& value, const Format& format)
{
	//	Delegate to sheet if present.  
	//	Offsets are relative to upper left cell range
	if (!d->sheet) return false;

	return d->sheet->write(d->top + row, d->left + col, value, format);
}

bool CellRange::writeString(int row, int col, const QString& value, const Format& format)
{
	//	Delegate to sheet if present.  
	//	Offsets are relative to upper left cell range
	if (!d->sheet) return false;

	return d->sheet->writeString(d->top + row, d->left + col, value, format);
}

Format CellRange::columnFormat( int col ) const
{
	if( !d->sheet ) return Format();

	return d->sheet->columnFormat( d->left + col );
}

bool CellRange::clear(int row, int col)
{
	//	Delegate to sheet if present.  
	//	Offsets are relative to upper left cell range
	if (!d->sheet) return false;

	return d->sheet->writeBlank(d->top + row, d->left + col);
}

QString CellRange::value2(int row, int col) const
{
	//	Delegate to sheet if present.  
	//	Offsets are relative to upper left cell range
	if (!d->sheet) return  QString();

	return d->sheet->value2(d->top + row, d->left + col);
}

Cell* CellRange::cellAt(int row, int col) const
{
	//	Delegate to sheet if present.  
	//	Offsets are relative to upper left cell range
	return d->sheet ? 
		d->sheet->cellAt(d->top+row, d->left + col) 
		: nullptr;
}

//	Search always occurs from upper left cell
CellRange CellRange::end(RangeSearch direction) const
{
	CellRange temp(*this);

	//	Handle down direction
	if (direction & xlDown)
	{
		//while (d->sheet->cellAt(++temp.d->bottom, d->left) && temp.d->bottom < XLSX_ROW_MAX);
		while (++temp.d->bottom && temp.d->bottom < XLSX_ROW_MAX)
		{
			//	Stop an null or empty cell
			const Cell* c = d->sheet->cellAt(temp.d->bottom, d->left);
			if (!c || c->value().isNull()) break;
		}

		//	Back off from empty cell
		--temp.d->bottom;
	}

	//	Handle up direction
	if (direction & xlUp)
	{
		//while (d->sheet->cellAt(--temp.d->top, d->left) && temp.d->top > 0);
		while (--temp.d->top && temp.d->top > 0)
		{
			//	Stop an null or empty cell
			const Cell* c = d->sheet->cellAt(temp.d->top, d->left);
			if (!c || c->value().isNull()) break;
		}

		//	Back off from empty cell
		++temp.d->top;
	}

	//	Handle right direction
	if (direction & xlRight)
	{
		while (++temp.d->right && temp.d->right < XLSX_COLUMN_MAX)
		{
			//	Stop an null or empty cell
			const Cell* c = d->sheet->cellAt(d->top, temp.d->right);
			if (!c || c->value().isNull()) break;
		}

		//	Back off from empty cell
		--temp.d->right;
	}

	//	Handle left direction
	if (direction & xlLeft)
	{
		//while (d->sheet->cellAt(d->top, --temp.d->left) && temp.d->left > 0)
		while (--temp.d->left && temp.d->left > 0)
		{
			//	Stop an null or empty cell
			const Cell* c = d->sheet->cellAt(d->top, temp.d->left);
			if (!c || c->value().isNull()) break;
		}

		//	Back off from empty cell
		++temp.d->left;
	}

	//	Use COW object
	return temp;
}

//	internal functions
void CellRange::setSheet(Worksheet *sheet)
{
	d->sheet = sheet;
}



QT_END_NAMESPACE_XLSX
