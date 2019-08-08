/****************************************************************************
** Copyright (c) 2018 David McRaven <david.mcraven@sabrixtax.com>
** All right reserved.
**
** Generic helper class to wrap QJsonObject for excel attributes
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
#pragma once

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

#include <QObject>
#include <QJsonObject>

//	Forward declarations
class QXmlStreamAttributes;

QT_BEGIN_NAMESPACE_XLSX

class Q_XLSX_EXPORT XlsxAttributes //: public QObject
{
	//Q_OBJECT

public:
	XlsxAttributes();
	~XlsxAttributes();
	XlsxAttributes(const XlsxAttributes&  other);
	XlsxAttributes(const XlsxAttributes&& other);

	qint32 count() const;
	bool toBool( const QString &key ) const;
	void setBool( const QString &key, bool flag );
	qint32 toInt( const QString &key ) const;
	void setInt( const QString &key, qint32 value );
	const QString toString( const QString &key ) const;
	void setString( const QString &key, const QString &value );

	QJsonObject::iterator insert( const QString &key, const QString &value );
	void remove( const QString &key );

	void readXMLAttributes( const QXmlStreamAttributes& attrs );
	const QXmlStreamAttributes writeXMLAttributes() const;

private:
	QJsonObject		_data;
};

QT_END_NAMESPACE_XLSX
