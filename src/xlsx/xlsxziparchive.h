#ifndef QXLSX_XLSXZIPARCHIVE_H
#define QXLSX_XLSXZIPARCHIVE_H

#include "xlsxglobal.h"

#include <QString>
#include <QByteArray>
#include <QStringList>

//  Forward declarations
class QIODevice;

QT_BEGIN_NAMESPACE_XLSX

//  Library forward declarations
//  Would use scoped or shareddata pointer, but it requires including private header
//  include in public include file (bad).
class XlsxZipArchivePrivate;

class XLSX_AUTOTEST_EXPORT XlsxZipArchive
{
public:
	enum class Status
	{
		NoError,
		FileWriteError,
		FileOpenError,
		FilePermissionsError,
		FileError
	};

    explicit XlsxZipArchive(const QString& archive);
    explicit XlsxZipArchive( QIODevice* file);

	XlsxZipArchive() = delete;
	~XlsxZipArchive();

    //	No copying
    XlsxZipArchive( const XlsxZipArchive& ) = delete;
    XlsxZipArchive& operator=( const XlsxZipArchive& ) = delete;
    //	Moving OK
    XlsxZipArchive( XlsxZipArchive&& ) = default;
    XlsxZipArchive& operator=( XlsxZipArchive&& ) = default;

	//	Reading operations
	bool exists() const;
	QStringList filePaths() const;
	QByteArray fileData( const QString &fileName ) const;

	//void addFile( const QString& filePath, QIODevice* device );
	void addFile( const QString& filePath, const QByteArray& data );
	bool error() const;
	void close();

private:

	XlsxZipArchivePrivate* const d = nullptr;
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_XLSXZIPARCHIVE
