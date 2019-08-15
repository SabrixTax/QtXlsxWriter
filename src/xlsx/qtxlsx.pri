INCLUDEPATH += $$PWD
DEPENDPATH += $$PWD

QT += core gui
!build_xlsx_lib:DEFINES += XLSX_NO_LIB

HEADERS += \
    $$PWD/XlsxAttributes_p.h \
    $$PWD/xlsxabstractooxmlfile.h \
    $$PWD/xlsxabstractooxmlfile_p.h \
    $$PWD/xlsxabstractsheet.h \
    $$PWD/xlsxabstractsheet_p.h \
    $$PWD/xlsxcell.h \
    $$PWD/xlsxcell_p.h \
    $$PWD/xlsxcellformula.h \
    $$PWD/xlsxcellformula_p.h \
    $$PWD/xlsxcellrange.h \
    $$PWD/xlsxcellrange_p.h \
    $$PWD/xlsxcellreference.h \
    $$PWD/xlsxchart.h \
    $$PWD/xlsxchart_p.h \
    $$PWD/xlsxchartsheet.h \
    $$PWD/xlsxchartsheet_p.h \
    $$PWD/xlsxcolor_p.h \
    $$PWD/xlsxconditionalformatting.h \
    $$PWD/xlsxconditionalformatting_p.h \
    $$PWD/xlsxcontenttypes_p.h \
    $$PWD/xlsxdatavalidation.h \
    $$PWD/xlsxdatavalidation_p.h \
    $$PWD/xlsxdefinednames.h \
    $$PWD/xlsxdefinednames_p.h \
    $$PWD/xlsxdocpropsapp_p.h \
    $$PWD/xlsxdocpropscore_p.h \
    $$PWD/xlsxdocument.h \
    $$PWD/xlsxdocument_p.h \
    $$PWD/xlsxdrawing_p.h \
    $$PWD/xlsxdrawinganchor_p.h \
    $$PWD/xlsxformat.h \
    $$PWD/xlsxformat_p.h \
    $$PWD/xlsxglobal.h \
    $$PWD/xlsxmediafile_p.h \
    $$PWD/xlsxnumformatparser_p.h \
    $$PWD/xlsxrelationships_p.h \
    $$PWD/xlsxrichstring.h \
    $$PWD/xlsxrichstring_p.h \
    $$PWD/xlsxsharedstrings_p.h \
    $$PWD/xlsxsimpleooxmlfile_p.h \
    $$PWD/xlsxstyles_p.h \
    $$PWD/xlsxtheme_p.h \
    $$PWD/xlsxutility_p.h \
    $$PWD/xlsxworkbook.h \
    $$PWD/xlsxworkbook_p.h \
    $$PWD/xlsxworksheet.h \
    $$PWD/xlsxworksheet_p.h \
    $$PWD/xlsxziparchive.h \
    $$PWD/xlsxziparchive_p.h \
    $$PWD/zip_file.h

SOURCES += \
    $$PWD/XlsxAttributes.cpp \
    $$PWD/xlsxabstractooxmlfile.cpp \
    $$PWD/xlsxabstractsheet.cpp \
    $$PWD/xlsxcell.cpp \
    $$PWD/xlsxcellformula.cpp \
    $$PWD/xlsxcellrange.cpp \
    $$PWD/xlsxcellreference.cpp \
    $$PWD/xlsxchart.cpp \
    $$PWD/xlsxchartsheet.cpp \
    $$PWD/xlsxcolor.cpp \
    $$PWD/xlsxconditionalformatting.cpp \
    $$PWD/xlsxcontenttypes.cpp \
    $$PWD/xlsxdatavalidation.cpp \
    $$PWD/xlsxdefinednames.cpp \
    $$PWD/xlsxdocpropsapp.cpp \
    $$PWD/xlsxdocpropscore.cpp \
    $$PWD/xlsxdocument.cpp \
    $$PWD/xlsxdrawing.cpp \
    $$PWD/xlsxdrawinganchor.cpp \
    $$PWD/xlsxformat.cpp \
    $$PWD/xlsxmediafile.cpp \
    $$PWD/xlsxnumformatparser.cpp \
    $$PWD/xlsxrelationships.cpp \
    $$PWD/xlsxrichstring.cpp \
    $$PWD/xlsxsharedstrings.cpp \
    $$PWD/xlsxsimpleooxmlfile.cpp \
    $$PWD/xlsxstyles.cpp \
    $$PWD/xlsxtheme.cpp \
    $$PWD/xlsxutility.cpp \
    $$PWD/xlsxworkbook.cpp \
    $$PWD/xlsxworksheet.cpp \
    $$PWD/xlsxziparchive.cpp
