#include "xlsxziparchive.h"
#include "xlsxziparchive_p.h"

//	Standard includes
#include <exception>

//	Qt includes
#include <QDebug>
#include <QFile>
#include <QString>

QT_BEGIN_NAMESPACE_XLSX

XlsxZipArchivePrivate::XlsxZipArchivePrivate()
	: archiver( new zip_file() )
{}

XlsxZipArchivePrivate::~XlsxZipArchivePrivate()
{
	//	If stream is open, close and delete it now.
	//if( qtDevice != nullptr )
	//{
	//	qtDevice->close();
		qtDevice = nullptr;
	//}
}

bool XlsxZipArchivePrivate::loadFromFileName( const QString &archive )
{
	try
	{
		//  Read zip file from hard disk
		archiver->load( archive.toStdString() );

		initRead();
	}
	catch( std::exception& e )
	{
		qDebug() << e.what();
		return false;
	}

	return true;
}

bool XlsxZipArchivePrivate::loadFromFile( QIODevice* device )
{
	//	If device is closed, this is an error. Just return.
	if( !device->isOpen() ) return false;

	//  Copy data to std::vector for zip library
	try
	{
		//	See if file is opened for reading
		QByteArray data = qtDevice->readAll();
		std::vector<unsigned char> zipData( data.begin(), data.end() );

		//  Read zip file from hard disk
		archiver->load( zipData );

		initRead();
	}
	catch( std::exception& e )
	{
		qDebug() << e.what();
		return false;
	}

	return true;
}

//  Get all file info
void XlsxZipArchivePrivate::initRead()
{
	//	File opened for reading
	isWrite = false;

	//  Loop through files in zip file
	//  Convert from std vector to QStringList
	for( const auto& fi : archiver->namelist() )
	{
		//	Convert from std::string to QT string
		filePaths.append( fi.c_str() );
	}
}

//  Initialize for for writing
void XlsxZipArchivePrivate::initWrite()
{
	//	File opened for reading
	isWrite = true;

	//	Get file name for later save
	fileNameFromHandle();
}

void XlsxZipArchivePrivate::fileNameFromHandle()
{
	QFile *f = qobject_cast<QFile*> ( qtDevice );
	
	if( f == nullptr )
		archiveName.clear();
	else
		archiveName = f->fileName().toStdString();
}

//  Declared private. Must associate with file on construction.
XlsxZipArchive::XlsxZipArchive(const QString &archive )
   : d( new XlsxZipArchivePrivate() )
{
	assert( d != nullptr );
	d->archiveName = archive.toStdString();
}

XlsxZipArchive::XlsxZipArchive(QIODevice* device)
	: d( new XlsxZipArchivePrivate() )
{
	assert( d != nullptr );
	assert( device->isOpen() );
	if( !device->isOpen() ) return;

	//	Looks like it's open. Cache for later
	d->qtDevice = device;

	//	See if file is opened for reading
	if( d->qtDevice->isReadable() )
	{
		d->loadFromFile( device );
	}
	else if( d->qtDevice->isWritable() )
	{
		//	Initialie for writing
		d->initWrite();
	}
}

XlsxZipArchive::~XlsxZipArchive()
{
   delete d;
}

QStringList XlsxZipArchive::filePaths() const
{
	return d->filePaths;
}

QByteArray XlsxZipArchive::fileData( const QString &fileName ) const
{
	//	Convert string returned to QByteArray
	const auto data = d->archiver->read( fileName.toStdString() );
	QByteArray temp( data.c_str(), static_cast<qint32>( data.length() ) );
	return temp;
}

//	Write operations
bool XlsxZipArchive::error() const
{
	//	Perform integrety check
	return !d->qtDevice->isOpen() || !d->qtDevice->isWritable();
}

/*
void XlsxZipArchive::addFile( const QString& filePath, QIODevice* device )
{
	Q_ASSERT( device );

	d->qtDevice = device;
	QFile *f = qobject_cast<QFile*> ( d->qtDevice );
	const std::string fileName = f->fileName().toStdString();

	QIODevice::OpenMode mode = device->openMode();
	bool opened = false;
	if( ( mode & QIODevice::ReadOnly ) == 0 )
	{
		opened = true;
		if( !device->open( QIODevice::ReadOnly ) )
		{
			d->status = Status::FileOpenError;
			return;
		}
	}

	d->archiver->writestr( fileName, device->readAll().toStdString() );
}
*/
void XlsxZipArchive::addFile( const QString &filePath, const QByteArray& data )
{
	d->archiver->writestr( filePath.toStdString(), data.toStdString() );
}

bool XlsxZipArchive::exists() const
{
	QFile *f = qobject_cast<QFile*> ( d->qtDevice );
	if( f == nullptr ) return true;
	
	return f->exists();
}


void XlsxZipArchive::close()
{
	//	This flushes contents to disk
	d->archiver->save( d->archiveName );

	//	Reset for next run, if any.
	d->archiver->reset();

	//	Our QT file handle is now not valid.
	d->qtDevice = nullptr;
}

QT_END_NAMESPACE_XLSX
