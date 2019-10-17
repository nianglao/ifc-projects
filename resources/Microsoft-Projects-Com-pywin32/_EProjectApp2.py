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

class _EProjectApp2:
	CLSID = CLSID_Sink = IID('{5066D7C4-1ED7-48C4-ACE7-299E109D368C}')
	coclass_clsid = IID('{36D27C48-A1E8-11D3-BA55-00C04F72F325}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		1073741837 : "OnProjectBeforeResourceNew2",
		       17 : "OnWindowBeforeViewChange",
		       11 : "OnProjectBeforeAssignmentChange",
		1073741830 : "OnProjectBeforeTaskDelete2",
		       31 : "OnApplicationBeforeClose",
		       29 : "OnProjectBeforeSaveBaseline",
		       43 : "OnConnectionStatusChanged",
		1073741832 : "OnProjectBeforeAssignmentDelete2",
		       12 : "OnProjectBeforeTaskNew",
		       22 : "OnWindowSidepaneTaskChange",
		       25 : "OnProjectAfterSave",
		        9 : "OnProjectBeforeTaskChange",
		1073741831 : "OnProjectBeforeResourceDelete2",
		       21 : "OnWindowSidepaneDisplayChange",
		       33 : "OnAfterCubeBuilt",
		        8 : "OnProjectBeforeAssignmentDelete",
		        4 : "OnProjectBeforePrint",
		       23 : "OnWorkpaneDisplayChange",
		       27 : "OnProjectResourceNew",
		1073741826 : "OnProjectBeforeClose2",
		1073741834 : "OnProjectBeforeResourceChange2",
		       28 : "OnProjectAssignmentNew",
		        1 : "OnNewProject",
		1073741828 : "OnProjectBeforePrint2",
		1073741838 : "OnProjectBeforeAssignmentNew2",
		       35 : "OnJobStart",
		       39 : "OnProjectBeforePublish",
		        2 : "OnProjectBeforeClose",
		       18 : "OnWindowViewChange",
		       24 : "OnLoadWebPage",
		       41 : "OnSecondaryViewChange",
		        3 : "OnProjectBeforeSave",
		        6 : "OnProjectBeforeTaskDelete",
		       10 : "OnProjectBeforeResourceChange",
		1073741833 : "OnProjectBeforeTaskChange2",
		       40 : "OnPaneActivate",
		       14 : "OnProjectBeforeAssignmentNew",
		       20 : "OnWindowDeactivate",
		       34 : "OnLoadWebPane",
		       13 : "OnProjectBeforeResourceNew",
		       37 : "OnSaveStartingToServer",
		       15 : "OnWindowGoalAreaChange",
		        5 : "OnProjectCalculate",
		        7 : "OnProjectBeforeResourceDelete",
		       32 : "OnUndoOrRedo",
		1073741836 : "OnProjectBeforeTaskNew2",
		       36 : "OnJobCompleted",
		       38 : "OnSaveCompletedToServer",
		       30 : "OnProjectBeforeClearBaseline",
		1073741827 : "OnProjectBeforeSave2",
		       42 : "OnIsFunctionalitySupported",
		       26 : "OnProjectTaskNew",
		       19 : "OnWindowActivate",
		1073741835 : "OnProjectBeforeAssignmentChange2",
		       16 : "OnWindowSelectionChange",
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
#	def OnProjectBeforeResourceNew2(self, pj=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnWindowBeforeViewChange(self, Window=defaultNamedNotOptArg, prevView=defaultNamedNotOptArg, newView=defaultNamedNotOptArg, projectHasViewWindow=defaultNamedNotOptArg
#			, Info=defaultNamedNotOptArg):
#	def OnProjectBeforeAssignmentChange(self, asg=defaultNamedNotOptArg, Field=defaultNamedNotOptArg, NewVal=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeTaskDelete2(self, tsk=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnApplicationBeforeClose(self, Info=defaultNamedNotOptArg):
#	def OnProjectBeforeSaveBaseline(self, pj=defaultNamedNotOptArg, Interim=defaultNamedNotOptArg, bl=defaultNamedNotOptArg, InterimCopy=defaultNamedNotOptArg
#			, InterimInto=defaultNamedNotOptArg, AllTasks=defaultNamedNotOptArg, RollupToSummaryTasks=defaultNamedNotOptArg, RollupFromSubtasks=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnConnectionStatusChanged(self, online=defaultNamedNotOptArg):
#	def OnProjectBeforeAssignmentDelete2(self, asg=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnProjectBeforeTaskNew(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWindowSidepaneTaskChange(self, Window=defaultNamedNotOptArg, ID=defaultNamedNotOptArg, IsGoalArea=defaultNamedNotOptArg):
#	def OnProjectAfterSave(self):
#	def OnProjectBeforeTaskChange(self, tsk=defaultNamedNotOptArg, Field=defaultNamedNotOptArg, NewVal=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeResourceDelete2(self, res=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnWindowSidepaneDisplayChange(self, Window=defaultNamedNotOptArg, Close=defaultNamedNotOptArg):
#	def OnAfterCubeBuilt(self, CubeFileName=defaultNamedNotOptArg):
#	def OnProjectBeforeAssignmentDelete(self, asg=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforePrint(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWorkpaneDisplayChange(self, DisplayState=defaultNamedNotOptArg):
#	def OnProjectResourceNew(self, pj=defaultNamedNotOptArg, ID=defaultNamedNotOptArg):
#	def OnProjectBeforeClose2(self, pj=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnProjectBeforeResourceChange2(self, res=defaultNamedNotOptArg, Field=defaultNamedNotOptArg, NewVal=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnProjectAssignmentNew(self, pj=defaultNamedNotOptArg, ID=defaultNamedNotOptArg):
#	def OnNewProject(self, pj=defaultNamedNotOptArg):
#	def OnProjectBeforePrint2(self, pj=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnProjectBeforeAssignmentNew2(self, pj=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnJobStart(self, bstrName=defaultNamedNotOptArg, bstrprojGuid=defaultNamedNotOptArg, bstrjobGuid=defaultNamedNotOptArg, jobType=defaultNamedNotOptArg
#			, lResult=defaultNamedNotOptArg):
#	def OnProjectBeforePublish(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeClose(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWindowViewChange(self, Window=defaultNamedNotOptArg, prevView=defaultNamedNotOptArg, newView=defaultNamedNotOptArg, success=defaultNamedNotOptArg):
#	def OnLoadWebPage(self, Window=defaultNamedNotOptArg, TargetPage=defaultNamedNotOptArg):
#	def OnSecondaryViewChange(self, Window=defaultNamedNotOptArg, prevView=defaultNamedNotOptArg, newView=defaultNamedNotOptArg, success=defaultNamedNotOptArg):
#	def OnProjectBeforeSave(self, pj=defaultNamedNotOptArg, SaveAsUi=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeTaskDelete(self, tsk=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeResourceChange(self, res=defaultNamedNotOptArg, Field=defaultNamedNotOptArg, NewVal=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnProjectBeforeTaskChange2(self, tsk=defaultNamedNotOptArg, Field=defaultNamedNotOptArg, NewVal=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnPaneActivate(self):
#	def OnProjectBeforeAssignmentNew(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWindowDeactivate(self, deactivatedWindow=defaultNamedNotOptArg):
#	def OnLoadWebPane(self, Window=defaultNamedNotOptArg, TargetPage=defaultNamedNotOptArg):
#	def OnProjectBeforeResourceNew(self, pj=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSaveStartingToServer(self, bstrName=defaultNamedNotOptArg, bstrprojGuid=defaultNamedNotOptArg):
#	def OnWindowGoalAreaChange(self, Window=defaultNamedNotOptArg, goalArea=defaultNamedNotOptArg):
#	def OnProjectCalculate(self, pj=defaultNamedNotOptArg):
#	def OnProjectBeforeResourceDelete(self, res=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnUndoOrRedo(self, bstrLabel=defaultNamedNotOptArg, bstrGUID=defaultNamedNotOptArg, fUndo=defaultNamedNotOptArg):
#	def OnProjectBeforeTaskNew2(self, pj=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnJobCompleted(self, bstrName=defaultNamedNotOptArg, bstrprojGuid=defaultNamedNotOptArg, bstrjobGuid=defaultNamedNotOptArg, jobType=defaultNamedNotOptArg
#			, lResult=defaultNamedNotOptArg):
#	def OnSaveCompletedToServer(self, bstrName=defaultNamedNotOptArg, bstrprojGuid=defaultNamedNotOptArg):
#	def OnProjectBeforeClearBaseline(self, pj=defaultNamedNotOptArg, Interim=defaultNamedNotOptArg, bl=defaultNamedNotOptArg, InterimFrom=defaultNamedNotOptArg
#			, AllTasks=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnProjectBeforeSave2(self, pj=defaultNamedNotOptArg, SaveAsUi=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnIsFunctionalitySupported(self, bstrFunctionality=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnProjectTaskNew(self, pj=defaultNamedNotOptArg, ID=defaultNamedNotOptArg):
#	def OnWindowActivate(self, activatedWindow=defaultNamedNotOptArg):
#	def OnProjectBeforeAssignmentChange2(self, asg=defaultNamedNotOptArg, Field=defaultNamedNotOptArg, NewVal=defaultNamedNotOptArg, Info=defaultNamedNotOptArg):
#	def OnWindowSelectionChange(self, Window=defaultNamedNotOptArg, sel=defaultNamedNotOptArg, selType=defaultNamedNotOptArg):


win32com.client.CLSIDToClass.RegisterCLSID( "{5066D7C4-1ED7-48C4-ACE7-299E109D368C}", _EProjectApp2 )
