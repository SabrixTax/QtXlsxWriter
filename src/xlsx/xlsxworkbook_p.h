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
#ifndef XLSXWORKBOOK_P_H
#define XLSXWORKBOOK_P_H

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

#include "xlsxworkbook.h"
#include "xlsxabstractooxmlfile_p.h"
#include "xlsxtheme_p.h"
#include "xlsxsimpleooxmlfile_p.h"
#include "xlsxrelationships_p.h"

//	Dave added
#include "xlsxdefinednames.h"
#include "XlsxAttributes_p.h"
//	Dave end

#include <QSharedPointer>
#include <QPair>
#include <QStringList>

namespace QXlsx {

/*
struct XlsxDefineNameData
{
    XlsxDefineNameData()
        :sheetId(-1)
    {}
    XlsxDefineNameData(const QString &name, const QString &formula, const QString &comment, int sheetId=-1)
        :name(name), formula(formula), comment(comment), sheetId(sheetId)
    {
	
    }
    QString name;
    QString formula;
    QString comment;
    //using internal sheetId, instead of the localSheetId(order in the workbook)
    int sheetId;
};
*/
class WorkbookPrivate : public AbstractOOXmlFilePrivate
{
    Q_DECLARE_PUBLIC(Workbook)
public:
    WorkbookPrivate(Workbook *q, Workbook::CreateFlag flag);

    QSharedPointer<SharedStrings> sharedStrings;
    QList<QSharedPointer<AbstractSheet> > sheets;
    QList<QSharedPointer<SimpleOOXmlFile> > externalLinks;
    QStringList sheetNames;
    QSharedPointer<Styles> styles;
    QSharedPointer<Theme> theme;
    QList<QSharedPointer<MediaFile> > mediaFiles;
    QList<QSharedPointer<Chart> > chartFiles;
    //QList<XlsxDefineNameData> definedNamesList;
		
	//	Lookup up sheets based on index
	//QVector<QSharedPointer<AbstractSheet> > sheetsIndex;
	//	Lookup based on sheet name
	//QMap<QString,QSharedPointer<AbstractSheet> > sheetsMap;

	//	Dave added
	void saveXmlHeader(QXmlStreamWriter &writer) const;	
	void loadXmlHeader(QXmlStreamReader &reader);	

	//	Initialize values for new spreadsheet
	void initalizeBlank();
	//	Dave end

    bool strings_to_numbers_enabled;	//					
    bool strings_to_hyperlinks_enabled;
    bool html_to_richstring_enabled;
    QString defaultDateFormat;

//	Replaced with workbookView below
    //int x_window;
    //int y_window;
    //int window_width;
    //int window_height;
	//int firstsheet;
	//int activesheetIndex;
	
    int table_count;

    //Used to generate new sheet name and id
    int last_worksheet_index;
    int last_chartsheet_index;
    int last_sheet_id;

	//	Dave-Significantly expand range logic
	QScopedPointer<DefinedNames>	definedNames;

	//	Dave-Replace hard coded values with variables
	QXmlStreamNamespaceDeclarations namespaceDeclarations;
	QXmlStreamAttributes			headerAttributes;
	QString							headerMamespace;

	//	Dave - File data
	//	Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference1
	//	Page 1565-1566
	//QXmlStreamAttributes			fileVersion;
	XlsxAttributes					fileVersion;

	//	Dave - workbookPr data
	//	Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference1
	//	Page 1576-1580 - Note, only one field picked up.  Many left to migrate
	XlsxAttributes					workbookPr;

	//	Dave - calcPr data
	//	Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference1
	//	Page 1543-1547 - Note, only one field picked up.  Many left to migrate
	XlsxAttributes					calcPr;

	//	Dave - bookviews data
	//	Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference1
	//	Page 1587-1589 - Note, only one field picked up.  Many left to migrate
	//QList<XlsxAttributes>			workbookView;
	XlsxAttributes					workbookView;
};

}

#endif // XLSXWORKBOOK_P_H
