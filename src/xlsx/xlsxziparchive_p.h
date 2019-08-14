/****************************************************************************
** Copyright (c) 2019 David McRaven <david.mcraven@sabrixtax.com>
** All right reserved.
**
** Replace private implementation of Zip with public version. This will solve
** needing to change private header references with every system QT upgrade.
** New header is zip_file.h. See file for copyright info.
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
#ifndef QXLSX_XLSXZIPARCHIVE_P_H
#define QXLSX_XLSXZIPARCHIVE_P_H

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

//	Standard library
#include <memory>	//	Unique_ptr
#include <string>

//	QT classes
#include <QStringList>

#include "xlsxglobal.h"
#include "xlsxziparchive.h"
#include "zip_file.h"

using namespace miniz_cpp;

class QIODevice;
class QStringList;

QT_BEGIN_NAMESPACE_XLSX


class XlsxZipArchivePrivate
{
public:
	XlsxZipArchivePrivate();
    ~XlsxZipArchivePrivate();

    //  No copying
	XlsxZipArchivePrivate( const XlsxZipArchivePrivate& ) = delete;
	XlsxZipArchivePrivate& operator=( const XlsxZipArchivePrivate& ) = delete;
    //	Moving OK
	XlsxZipArchivePrivate( XlsxZipArchivePrivate&& ) = default;
	XlsxZipArchivePrivate& operator=( XlsxZipArchivePrivate&& ) = default;

	//  Load reader with stream
	bool loadFromFileName( const QString &archive );
	bool loadFromFile( QIODevice* device );

	void initWrite();

	//	Used for write functions
	bool						isWrite = false;
	QIODevice*					qtDevice = nullptr;

    std::unique_ptr<zip_file>	archiver;
	std::string					archiveName;
	QStringList                 filePaths;
	XlsxZipArchive::Status		status = XlsxZipArchive::Status::NoError;

private:
	//  initialize class for use
	void initRead();
	void fileNameFromHandle();
};

QT_END_NAMESPACE_XLSX


#endif // QXLSX_XLSXZIPARCHIVE_P_H
