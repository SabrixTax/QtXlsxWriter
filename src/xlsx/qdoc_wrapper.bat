@echo off
SetLocal EnableDelayedExpansion
(set QT_VERSION=0.3.0)
(set QT_VER=0.3)
(set QT_VERSION_TAG=030)
(set QT_INSTALL_DOCS=C:/QT/Docs/Qt-5.10.1)
(set BUILDDIR=C:/Save/Projects/QT/QtXlsxWriter/src/xlsx)
C:\QT\5.10.1\MSVC2017_64\bin\qdoc.exe %*
EndLocal
