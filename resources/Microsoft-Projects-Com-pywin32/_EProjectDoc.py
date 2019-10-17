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

class _EProjectDoc:
	CLSID = CLSID_Sink = IID('{F81DD3C0-5089-11CF-A49D-00AA00574C74}')
	coclass_clsid = IID('{1019A320-508A-11CF-A49D-00AA00574C74}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		        7 : "OnActivate",
		        5 : "OnCalculate",
		        3 : "OnBeforeSave",
		        2 : "OnBeforeClose",
		        8 : "OnDeactivate",
		        4 : "OnBeforePrint",
		        1 : "OnOpen",
		        6 : "OnChange",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnActivate(self, pj=defaultNamedNotOptArg):
#	def OnCalculate(self, pj=defaultNamedNotOptArg):
#	def OnBeforeSave(self, pj=defaultNamedNotOptArg):
#	def OnBeforeClose(self, pj=defaultNamedNotOptArg):
#	def OnDeactivate(self, pj=defaultNamedNotOptArg):
#	def OnBeforePrint(self, pj=defaultNamedNotOptArg):
#	def OnOpen(self, pj=defaultNamedNotOptArg):
#	def OnChange(self, pj=defaultNamedNotOptArg):


win32com.client.CLSIDToClass.RegisterCLSID( "{F81DD3C0-5089-11CF-A49D-00AA00574C74}", _EProjectDoc )
