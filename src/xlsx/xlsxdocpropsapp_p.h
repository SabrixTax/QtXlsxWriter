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
#ifndef XLSXDOCPROPSAPP_H
#define XLSXDOCPROPSAPP_H

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
#include "xlsxabstractooxmlfile.h"
#include <QList>
#include <QPair>
#include <QStringList>
#include <QMap>

class QIODevice;

namespace QXlsx {

class XLSX_AUTOTEST_EXPORT DocPropsApp : public AbstractOOXmlFile
{
public:
    DocPropsApp(CreateFlag flag);
    
    void addPartTitle(const QString &title);
    void addHeadingPair(const QString &name, int value);

    bool setProperty(const QString &name, const QString &value);
    QString property(const QString &name) const;
    QStringList propertyNames() const;

    void saveToXmlFile(QIODevice *device) const;
    bool loadFromXmlFile(QIODevice *device);

	//	Dave-Add accessors to new fields
	QString application() const;
	void setApplication(const QString& value);
	QString docSecurity() const;
	void setDocSecurity(const QString& value);
	QString scaleCrop() const;
	void setScaleCrop(const QString& value);
	QString linksUpToDate() const;
	void setLinksUpToDate(const QString& value);
	QString sharedDoc() const;
	void setSharedDoc(const QString& value);
	QString hyperlinksChanged() const;
	void setHyperlinksChanged(const QString& value);
	QString appVersion() const;
	void setAppVersion(const QString& value);

private:
    QStringList m_titlesOfPartsList;
    QList<QPair<QString, int> > m_headingPairsList;
    QMap<QString, QString> m_properties;

	//	Dave-Replace hard coded values with cached values from open file
	QString		m_application;
	QString		m_docSecurity;
	QString		m_scaleCrop;
	QString		m_linksUpToDate;
	QString		m_sharedDoc;
	QString		m_hyperlinksChanged;
	QString		m_appVersion;
};

}
#endif // XLSXDOCPROPSAPP_H
