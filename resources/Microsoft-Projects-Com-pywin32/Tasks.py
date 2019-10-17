# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 2.7.16 (v2.7.16:413a49145e, Mar  4 2019, 01:37:19) [MSC v.1500 64 bit (AMD64)]
# From type library '{A7107640-94DF-1068-855E-00DD01075445}'
# On Thu Oct 17 14:23:50 2019
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

from win32com.client import DispatchBaseClass
class Tasks(DispatchBaseClass):
	CLSID = IID('{000C0C40-0000-0000-C000-000000000046}')
	coclass_clsid = None

	# Result is of type Task
	def Add(self, Name=defaultNamedOptArg, Before=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(65526, LCID, 1, (9, 0), ((12, 17), (12, 17)),Name
			, Before)
		if ret is not None:
			ret = Dispatch(ret, u'Add', '{000C0C3F-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Task
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, u'Item', '{000C0C3F-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Task
	# The method UniqueID is actually a property, but must be used as a method to correctly pass the arguments
	def UniqueID(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(65521, LCID, 2, (9, 0), ((3, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, u'UniqueID', '{000C0C3F-0000-0000-C000-000000000046}')
		return ret

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (65524, 2, (13, 0), (), "Application", '{36D27C48-A1E8-11D3-BA55-00C04F72F325}'),
		"Count": (65525, 2, (3, 0), (), "Count", None),
		# Method 'Parent' returns object of type 'Project'
		"Parent": (65523, 2, (13, 0), (), "Parent", '{1019A320-508A-11CF-A49D-00AA00574C74}'),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{000C0C3F-0000-0000-C000-000000000046}')
		return ret

	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,2,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, '{000C0C3F-0000-0000-C000-000000000046}')
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(65525, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

win32com.client.CLSIDToClass.RegisterCLSID( "{000C0C40-0000-0000-C000-000000000046}", Tasks )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 2.7.16 (v2.7.16:413a49145e, Mar  4 2019, 01:37:19) [MSC v.1500 64 bit (AMD64)]
# From type library '{A7107640-94DF-1068-855E-00DD01075445}'
# On Thu Oct 17 14:23:50 2019
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

Tasks_vtables_dispatch_ = 1
Tasks_vtables_ = [
	(( u'Application' , u'retval' , ), 65524, (65524, (), [ (16397, 10, None, "IID('{36D27C48-A1E8-11D3-BA55-00C04F72F325}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( u'Parent' , u'retval' , ), 65523, (65523, (), [ (16397, 10, None, "IID('{1019A320-508A-11CF-A49D-00AA00574C74}')") , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( u'Count' , u'retval' , ), 65525, (65525, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( u'UniqueID' , u'Index' , u'retval' , ), 65521, (65521, (), [ (3, 1, None, None) , 
			(16393, 10, None, "IID('{000C0C3F-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( u'Add' , u'Name' , u'Before' , u'retval' , ), 65526, (65526, (), [ 
			(12, 17, None, None) , (12, 17, None, None) , (16393, 10, None, "IID('{000C0C3F-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 88 , (3, 0, None, None) , 0 , )),
	(( u'Item' , u'Index' , u'retval' , ), 0, (0, (), [ (12, 1, None, None) , 
			(16393, 10, None, "IID('{000C0C3F-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( u'_NewEnum' , u'ppEnumVar' , ), -4, (-4, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 1024 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{000C0C40-0000-0000-C000-000000000046}", Tasks )
