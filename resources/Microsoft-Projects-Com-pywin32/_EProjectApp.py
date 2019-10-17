# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 2.7.16 (v2.7.16:413a49145e, Mar  4 2019, 01:37:19) [MSC v.1500 64 bit (AMD64)]
# From type library '{A7107640-94DF-1068-855E-00DD01075445}'
# On Thu Oct 17 14:14:47 2019
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

class _EProjectApp:
	CLSID = CLSID_Sink = IID('{7B7597D0-0C9F-11D0-8C43-00A0C904DCD4}')
	coclass_clsid = IID('{36D27C48-A1E8-11D3-BA55-00C04F72F325}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		        7 : "OnProjectBeforeResourceDelete",
		        8 : "OnProjectBeforeAssignmentDelete",
		       11 : "OnProjectBeforeAssignmentChange",
		        3 : "OnProjectBeforeSave",
		       10 : "OnProjectBeforeResourceChange",
		        1 : "OnNewProject",
		        4 : "OnProjectBeforePrint",
		       14 : "OnProjectBeforeAssignmentNew",
		       12 : "OnProjectBeforeTaskNew",
		        6 : "OnProjectBeforeTaskDelete",
		       13 : "OnProjectBeforeResourceNew",
		        9 : "OnProjectBeforeTaskChange",
		        2 : "OnProjectBeforeClose",
		        5 : "OnProjectCalculate",
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
#	def OnProjectBeforeResourceDelete(self, res=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeAssignmentDelete(self, asg=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeAssignmentChange(self, asg=defaultNamedNotOptArg, Field=defaultNamedNotOptArg, NewVal=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeSave(self, pj=defaultNamedNotOptArg, SaveAsUi=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeResourceChange(self, res=defaultNamedNotOptArg, Field=defaultNamedNotOptArg, NewVal=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnNewProject(self, pj=defaultNamedNotOptArg):
#	def OnProjectBeforePrint(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeAssignmentNew(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeTaskNew(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeTaskDelete(self, tsk=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeResourceNew(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeTaskChange(self, tsk=defaultNamedNotOptArg, Field=defaultNamedNotOptArg, NewVal=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeClose(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectCalculate(self, pj=defaultNamedNotOptArg):


win32com.client.CLSIDToClass.RegisterCLSID( "{7B7597D0-0C9F-11D0-8C43-00A0C904DCD4}", _EProjectApp )
