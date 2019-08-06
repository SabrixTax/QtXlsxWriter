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
#include "XlsxAttributes_p.h"

#include <QXmlStreamAttribute>
#include <QXmlStreamAttributes>

#include "xlsxglobal.h"

//	Cannot use QStreamAttributes since each attribute is read only
//	Cannot use QJsonValue since types could change (eg read string but process as bool or int)
//	if wrong type used for QJsonValue, empty value is returned.  Want similar behavior as this class
//	but using QVariants for value to allow changing of data types.
QT_BEGIN_NAMESPACE_XLSX

XlsxAttributes::XlsxAttributes()
	//: QObject(parent)
{
}

XlsxAttributes::XlsxAttributes(const XlsxAttributes& other) :
	_data( other._data )
{
}

XlsxAttributes::XlsxAttributes(const XlsxAttributes&& other) :
	_data( std::move( other._data ) )
{
}

XlsxAttributes::~XlsxAttributes()
{
}

qint32 XlsxAttributes::count() const
{	return _data.count(); }

//	For now, treat all attributes as string
QJsonObject::iterator XlsxAttributes::insert( const QString &key, const QString &value )
{
	return _data.insert( key, value );
}

void XlsxAttributes::remove( const QString &key )
{
	_data.remove( key );
}

bool XlsxAttributes::toBool( const QString &key ) const
{
	QJsonObject::const_iterator value = _data.find( key );
	
	if( value != _data.constEnd() )
		//	Excel uses 1 to represent true;
		return value.value().toString() == QStringLiteral( "1" );

	//	Missing record is a fail
	return false;
}

void XlsxAttributes::setBool( const QString &key, bool flag ) 
{
	//	Excel saves true and false as 1 and 0, respectively
	const QJsonValue newValue( flag ? QStringLiteral( "1" ) : QStringLiteral( "0" ) );

	//	Insert will overwrite existing value
	_data.insert( key, newValue );
}

qint32 XlsxAttributes::toInt( const QString &key ) const
{
	QJsonObject::const_iterator value = _data.find( key );

	if( value != _data.constEnd() )
		//	Convert string to integer and return;
		return value.value().toString().toInt();

	//	Missing record is treated as 0
	return 0;
}

void XlsxAttributes::setInt( const QString &key, qint32 value ) 
{
	//	need to convert to string and then QJsonValue object
	QString temp = QString("%1").arg(value);
	const QJsonValue newValue( temp );

	//	Insert will overwrite existing value
	_data.insert( key, newValue );
}

const QString XlsxAttributes::toString( const QString &key ) const
{
	QJsonObject::const_iterator value = _data.find( key );

	if( value != _data.constEnd() )
		//	Convert string to integer and return;
		return value.value().toString();

	//	Missing record is treated as empty Qstring
	return QString();
}

void XlsxAttributes::setString( const QString &key, const QString &value ) 
{
	//	need to convert to string and then QJsonValue object
	const QJsonValue newValue( value );

	//	Insert will overwrite existing value
	_data.insert( key, newValue );
}

void XlsxAttributes::readXMLAttributes( const QXmlStreamAttributes& attrs )
{
	//	Attributes are not nested.  Can read them in directly
	for( const QXmlStreamAttribute& attr : attrs )
	{
		//	All values are saved as strings
		_data.insert( attr.name().toString(), attr.value().toString() );
	}
}

const QXmlStreamAttributes XlsxAttributes::writeXMLAttributes() const
{
	//	Create temporary attribute container
	QXmlStreamAttributes attr;

	//	Attributes are not nested.  Can read them in directly
	for( const QString& key : _data.keys() )
	{
		QJsonValue value = _data.value( key );

		//	Everything saved as string in attributes
		attr.append(key, value.toString());
	}

	//	This will return a copy
	return attr;
}

QT_END_NAMESPACE_XLSX