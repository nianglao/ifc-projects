# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 2.7.16 (v2.7.16:413a49145e, Mar  4 2019, 01:37:19) [MSC v.1500 64 bit (AMD64)]
# From type library '{A7107640-94DF-1068-855E-00DD01075445}'
# On Thu Oct 17 14:22:34 2019
'Microsoft Project 14.0 Object Library'
makepy_version = '0.5.01'
python_version = 0x20710f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{A7107640-94DF-1068-855E-00DD01075445}')
MajorVersion = 4
MinorVersion = 7
LibraryFlags = 8
LCID = 0x0

from win32com.client import CoClassBaseClass
import sys
__import__('win32com.gen_py.A7107640-94DF-1068-855E-00DD01075445x0x4x7._EProjectDoc')
_EProjectDoc = sys.modules['win32com.gen_py.A7107640-94DF-1068-855E-00DD01075445x0x4x7._EProjectDoc']._EProjectDoc
__import__('win32com.gen_py.A7107640-94DF-1068-855E-00DD01075445x0x4x7._IProjectDoc')
_IProjectDoc = sys.modules['win32com.gen_py.A7107640-94DF-1068-855E-00DD01075445x0x4x7._IProjectDoc']._IProjectDoc
class Project(CoClassBaseClass): # A CoClass
	CLSID = IID('{1019A320-508A-11CF-A49D-00AA00574C74}')
	coclass_sources = [
		_EProjectDoc,
	]
	default_source = _EProjectDoc
	coclass_interfaces = [
		_IProjectDoc,
	]
	default_interface = _IProjectDoc

win32com.client.CLSIDToClass.RegisterCLSID( "{1019A320-508A-11CF-A49D-00AA00574C74}", Project )
