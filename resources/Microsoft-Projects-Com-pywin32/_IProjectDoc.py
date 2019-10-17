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

from win32com.client import DispatchBaseClass
class _IProjectDoc(DispatchBaseClass):
	CLSID = IID('{00020B00-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{1019A320-508A-11CF-A49D-00AA00574C74}')

	def Activate(self):
		return self._oleobj_.InvokeTypes(13, LCID, 1, (24, 0), (),)

	def AppendNotes(self, Value=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(79, LCID, 1, (24, 0), ((8, 0),),Value
			)

	# The method BaselineSavedDate is actually a property, but must be used as a method to correctly pass the arguments
	def BaselineSavedDate(self, Baseline=defaultNamedNotOptArg):
		return self._ApplyTypes_(6022, 2, (12, 0), ((3, 0),), u'BaselineSavedDate', None,Baseline
			)

	def CheckIn(self, SaveChanges=defaultNamedOptArg, Comment=defaultNamedOptArg, MakePublic=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(6053, LCID, 1, (24, 0), ((12, 16), (12, 16), (12, 16)),SaveChanges
			, Comment, MakePublic)

	def DeliverableAcceptChanges(self, DeliverableGuid=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(6075, LCID, 1, (11, 0), ((8, 0),),DeliverableGuid
			)

	def DeliverableCreate(self, DeliverableName=defaultNamedNotOptArg, DeliverableStartDate=defaultNamedNotOptArg, DeliverableFinishDate=defaultNamedNotOptArg, TaskGuid=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(6062, LCID, 1, (8, 0), ((8, 0), (12, 0), (12, 0), (8, 0)),DeliverableName
			, DeliverableStartDate, DeliverableFinishDate, TaskGuid)

	def DeliverableDelete(self, DeliverableGuid=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(6064, LCID, 1, (11, 0), ((8, 0),),DeliverableGuid
			)

	def DeliverableDependencyCreate(self, DeliverableGuid=defaultNamedNotOptArg, TaskGuid=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(6065, LCID, 1, (11, 0), ((8, 0), (8, 0)),DeliverableGuid
			, TaskGuid)

	def DeliverableDependencyDelete(self, DeliverableGuid=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(6066, LCID, 1, (11, 0), ((8, 0),),DeliverableGuid
			)

	def DeliverableLinkToProject(self, DeliverableGuid=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(6073, LCID, 1, (11, 0), ((8, 0),),DeliverableGuid
			)

	def DeliverableLinkToTask(self, DeliverableGuid=defaultNamedNotOptArg, TaskGuid=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(6072, LCID, 1, (11, 0), ((8, 0), (8, 0)),DeliverableGuid
			, TaskGuid)

	def DeliverableRefreshServerCache(self, DeliverableGuid=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(6067, LCID, 1, (11, 0), ((12, 16),),DeliverableGuid
			)

	def DeliverableUpdate(self, DeliverableGuid=defaultNamedNotOptArg, DeliverableName=defaultNamedNotOptArg, DeliverableStartDate=defaultNamedNotOptArg, DeliverableFinishDate=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(6063, LCID, 1, (11, 0), ((8, 0), (8, 0), (12, 0), (12, 0)),DeliverableGuid
			, DeliverableName, DeliverableStartDate, DeliverableFinishDate)

	def DeliverablesClearAll(self):
		return self._oleobj_.InvokeTypes(6074, LCID, 1, (11, 0), (),)

	def DeliverablesGetByProject(self, ProjectGuid=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(6077, LCID, 1, (9, 0), ((8, 0),),ProjectGuid
			)
		if ret is not None:
			ret = Dispatch(ret, u'DeliverablesGetByProject', None)
		return ret

	def DeliverablesGetProviderProjects(self):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(6076, LCID, 1, (8, 0), (),)

	def DeliverablesGetServerCachedXml(self):
		ret = self._oleobj_.InvokeTypes(6068, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, u'DeliverablesGetServerCachedXml', None)
		return ret

	def DeliverablesGetXml(self):
		ret = self._oleobj_.InvokeTypes(6069, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, u'DeliverablesGetXml', None)
		return ret

	def ExportAsFixedFormat(self, Filename=defaultNamedNotOptArg, FileType=0, IncludeDocumentProperties=True, IncludeDocumentMarkup=True
			, ArchiveFormat=False, FromDate=defaultNamedOptArg, ToDate=defaultNamedOptArg, FixedFormatExtClassPtr=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(6130, LCID, 1, (24, 0), ((8, 1), (3, 48), (11, 48), (11, 48), (11, 48), (12, 16), (12, 16), (12, 16)),Filename
			, FileType, IncludeDocumentProperties, IncludeDocumentMarkup, ArchiveFormat, FromDate
			, ToDate, FixedFormatExtClassPtr)

	def GetDisplayNameFromObjectMatchingID(self, ObjectType=defaultNamedNotOptArg, MatchingID=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(6060, LCID, 1, (8, 0), ((3, 0), (8, 0)),ObjectType
			, MatchingID)

	def GetObjectMatchingID(self, ObjectType=defaultNamedNotOptArg, ObjectName=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(6058, LCID, 1, (8, 0), ((3, 0), (8, 0)),ObjectType
			, ObjectName)

	def GetServerProjectGuid(self):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(6070, LCID, 1, (8, 0), (),)

	def GetTaskIndexByGuid(self, TaskGuid=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(6078, LCID, 1, (3, 0), ((8, 0),),TaskGuid
			)

	def GetWinprojURLs(self):
		return self._ApplyTypes_(6110, 1, (12, 0), (), u'GetWinprojURLs', None,)

	def ImportResourceErrorCount(self):
		return self._oleobj_.InvokeTypes(6116, LCID, 1, (3, 0), (),)

	def LevelClearDates(self):
		return self._oleobj_.InvokeTypes(78, LCID, 1, (24, 0), (),)

	def LocalResourceCount(self):
		return self._oleobj_.InvokeTypes(6114, LCID, 1, (3, 0), (),)

	def LocalResourceErrorCount(self):
		return self._oleobj_.InvokeTypes(6112, LCID, 1, (3, 0), (),)

	def MakeServerURLTrusted(self):
		return self._oleobj_.InvokeTypes(6037, LCID, 1, (24, 0), (),)

	def RSVDragSimulator(self, AssignmentToDrag=defaultNamedNotOptArg, DestinationResource=defaultNamedOptArg, DestinationTime=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(6127, LCID, 1, (24, 0), ((12, 1), (12, 16), (12, 16)),AssignmentToDrag
			, DestinationResource, DestinationTime)

	def ReadWssData(self, ProjectGuid=defaultNamedNotOptArg):
		return self._ApplyTypes_(6109, 1, (12, 0), ((8, 0),), u'ReadWssData', None,ProjectGuid
			)

	def ResourceCount(self):
		return self._oleobj_.InvokeTypes(6115, LCID, 1, (3, 0), (),)

	def ResourceErrorCount(self):
		return self._oleobj_.InvokeTypes(6113, LCID, 1, (3, 0), (),)

	def SaveAs(self, Name=defaultNamedNotOptArg, Format=0, Backup=defaultNamedOptArg, ReadOnly=defaultNamedOptArg
			, TaskInformation=defaultNamedOptArg, Filtered=defaultNamedOptArg, Table=defaultNamedOptArg, UserID=defaultNamedOptArg, DatabasePassWord=defaultNamedOptArg
			, FormatID=defaultNamedOptArg, Map=defaultNamedOptArg, ClearBaseline=defaultNamedOptArg, ClearActuals=defaultNamedOptArg, ClearResourceRates=defaultNamedOptArg
			, ClearFixedCosts=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(6003, LCID, 1, (24, 0), ((12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Format, Backup, ReadOnly, TaskInformation, Filtered
			, Table, UserID, DatabasePassWord, FormatID, Map
			, ClearBaseline, ClearActuals, ClearResourceRates, ClearFixedCosts)

	def SetCustomUI(self, CustomUIXML=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(6129, LCID, 1, (24, 0), ((8, 1),),CustomUIXML
			)

	def SetObjectMatchingID(self, ObjectType=defaultNamedNotOptArg, ObjectName=defaultNamedNotOptArg, MatchingID=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(6059, LCID, 1, (24, 0), ((3, 0), (8, 0), (8, 0)),ObjectType
			, ObjectName, MatchingID)

	_prop_map_get_ = {
		"AcceptNewExternalData": (88, 2, (11, 0), (), "AcceptNewExternalData", None),
		"ActualWork": (102, 2, (12, 0), (), "ActualWork", None),
		"AdministrativeProject": (6055, 2, (11, 0), (), "AdministrativeProject", None),
		"AllowTaskDelegation": (5129, 2, (11, 0), (), "AllowTaskDelegation", None),
		"AndMoveCompleted": (6043, 2, (11, 0), (), "AndMoveCompleted", None),
		"AndMoveRemaining": (6041, 2, (11, 0), (), "AndMoveRemaining", None),
		"AskForCompletedWork": (6000, 2, (3, 0), (), "AskForCompletedWork", None),
		"Author": (48, 2, (12, 0), (), "Author", None),
		"AutoAddResources": (5122, 2, (11, 0), (), "AutoAddResources", None),
		"AutoCalcCosts": (84, 2, (11, 0), (), "AutoCalcCosts", None),
		"AutoFilter": (66, 2, (11, 0), (), "AutoFilter", None),
		"AutoLinkTasks": (5114, 2, (11, 0), (), "AutoLinkTasks", None),
		"AutoSplitTasks": (5113, 2, (11, 0), (), "AutoSplitTasks", None),
		"AutoTrack": (5112, 2, (11, 0), (), "AutoTrack", None),
		"BaselineCost": (106, 2, (12, 0), (), "BaselineCost", None),
		"BaselineFinish": (144, 2, (12, 0), (), "BaselineFinish", None),
		"BaselineStart": (143, 2, (12, 0), (), "BaselineStart", None),
		"BaselineWork": (101, 2, (12, 0), (), "BaselineWork", None),
		"Comments": (14, 2, (12, 0), (), "Comments", None),
		"Company": (4, 2, (12, 0), (), "Company", None),
		"Contact": (212, 2, (8, 0), (), "Contact", None),
		"Cost1": (206, 2, (12, 0), (), "Cost1", None),
		"Cost2": (207, 2, (12, 0), (), "Cost2", None),
		"Cost3": (208, 2, (12, 0), (), "Cost3", None),
		"CurrencyCode": (5131, 2, (8, 0), (), "CurrencyCode", None),
		"CurrencyDigits": (5103, 2, (2, 0), (), "CurrencyDigits", None),
		"CurrencySymbol": (5101, 2, (8, 0), (), "CurrencySymbol", None),
		"CurrencySymbolPosition": (5102, 2, (3, 0), (), "CurrencySymbolPosition", None),
		"CurrentDate": (9, 2, (12, 0), (), "CurrentDate", None),
		"DatabaseProjectUniqueID": (6019, 2, (12, 0), (), "DatabaseProjectUniqueID", None),
		"DayLabelDisplay": (58, 2, (2, 0), (), "DayLabelDisplay", None),
		"DaysPerMonth": (5128, 2, (5, 0), (), "DaysPerMonth", None),
		"DefaultDurationUnits": (5108, 2, (3, 0), (), "DefaultDurationUnits", None),
		"DefaultEarnedValueMethod": (6044, 2, (3, 0), (), "DefaultEarnedValueMethod", None),
		"DefaultEffortDriven": (63, 2, (11, 0), (), "DefaultEffortDriven", None),
		"DefaultFinishTime": (5123, 2, (12, 0), (), "DefaultFinishTime", None),
		"DefaultFixedCostAccrual": (81, 2, (3, 0), (), "DefaultFixedCostAccrual", None),
		"DefaultResourceOvertimeRate": (5120, 2, (12, 0), (), "DefaultResourceOvertimeRate", None),
		"DefaultResourceStandardRate": (5119, 2, (12, 0), (), "DefaultResourceStandardRate", None),
		"DefaultStartTime": (5115, 2, (12, 0), (), "DefaultStartTime", None),
		"DefaultTaskType": (67, 2, (3, 0), (), "DefaultTaskType", None),
		"DefaultWorkUnits": (5109, 2, (3, 0), (), "DefaultWorkUnits", None),
		"Delay": (120, 2, (12, 0), (), "Delay", None),
		"DisplayProjectSummaryTask": (5121, 2, (11, 0), (), "DisplayProjectSummaryTask", None),
		"Duration1": (203, 2, (12, 0), (), "Duration1", None),
		"Duration2": (204, 2, (12, 0), (), "Duration2", None),
		"Duration3": (205, 2, (12, 0), (), "Duration3", None),
		"EarnedValueBaseline": (6045, 2, (3, 0), (), "EarnedValueBaseline", None),
		"EnterpriseActualsSynched": (6051, 2, (11, 0), (), "EnterpriseActualsSynched", None),
		"ExpandDatabaseTimephasedData": (6020, 2, (11, 0), (), "ExpandDatabaseTimephasedData", None),
		"Finish1": (153, 2, (12, 0), (), "Finish1", None),
		"Finish2": (156, 2, (12, 0), (), "Finish2", None),
		"Finish3": (159, 2, (12, 0), (), "Finish3", None),
		"Finish4": (162, 2, (12, 0), (), "Finish4", None),
		"Finish5": (165, 2, (12, 0), (), "Finish5", None),
		"FixedCost": (108, 2, (12, 0), (), "FixedCost", None),
		"FixedDuration": (134, 2, (12, 0), (), "FixedDuration", None),
		"Flag1": (172, 2, (12, 0), (), "Flag1", None),
		"Flag10": (181, 2, (12, 0), (), "Flag10", None),
		"Flag2": (173, 2, (12, 0), (), "Flag2", None),
		"Flag3": (174, 2, (12, 0), (), "Flag3", None),
		"Flag4": (175, 2, (12, 0), (), "Flag4", None),
		"Flag5": (176, 2, (12, 0), (), "Flag5", None),
		"Flag6": (177, 2, (12, 0), (), "Flag6", None),
		"Flag7": (178, 2, (12, 0), (), "Flag7", None),
		"Flag8": (179, 2, (12, 0), (), "Flag8", None),
		"Flag9": (180, 2, (12, 0), (), "Flag9", None),
		"FollowedHyperlinkColor": (98, 2, (3, 0), (), "FollowedHyperlinkColor", None),
		"FollowedHyperlinkColorEx": (6120, 2, (3, 0), (), "FollowedHyperlinkColorEx", None),
		"HideBar": (209, 2, (12, 0), (), "HideBar", None),
		"HonorConstraints": (73, 2, (11, 0), (), "HonorConstraints", None),
		"HourLabelDisplay": (57, 2, (2, 0), (), "HourLabelDisplay", None),
		"HoursPerDay": (5117, 2, (5, 0), (), "HoursPerDay", None),
		"HoursPerWeek": (5118, 2, (5, 0), (), "HoursPerWeek", None),
		"HyperlinkColor": (97, 2, (3, 0), (), "HyperlinkColor", None),
		"HyperlinkColorEx": (6119, 2, (3, 0), (), "HyperlinkColorEx", None),
		"IsTemplate": (6111, 2, (11, 0), (), "IsTemplate", None),
		"Keywords": (49, 2, (12, 0), (), "Keywords", None),
		"LevelEntireProject": (91, 2, (11, 0), (), "LevelEntireProject", None),
		"LevelFromDate": (76, 2, (12, 0), (), "LevelFromDate", None),
		"LevelToDate": (77, 2, (12, 0), (), "LevelToDate", None),
		"Manager": (5, 2, (12, 0), (), "Manager", None),
		"Marked": (171, 2, (12, 0), (), "Marked", None),
		"MinuteLabelDisplay": (56, 2, (2, 0), (), "MinuteLabelDisplay", None),
		"MonthLabelDisplay": (6015, 2, (2, 0), (), "MonthLabelDisplay", None),
		"MoveCompleted": (6040, 2, (11, 0), (), "MoveCompleted", None),
		"MoveRemaining": (6042, 2, (11, 0), (), "MoveRemaining", None),
		"MultipleCriticalPaths": (74, 2, (11, 0), (), "MultipleCriticalPaths", None),
		"Name": (0, 2, (8, 0), (), "Name", None),
		"NewTasksEstimated": (6005, 2, (11, 0), (), "NewTasksEstimated", None),
		"Notes": (115, 2, (8, 0), (), "Notes", None),
		"Number1": (187, 2, (5, 0), (), "Number1", None),
		"Number2": (188, 2, (5, 0), (), "Number2", None),
		"Number3": (189, 2, (5, 0), (), "Number3", None),
		"Number4": (190, 2, (5, 0), (), "Number4", None),
		"Number5": (191, 2, (5, 0), (), "Number5", None),
		"PercentWorkComplete": (133, 2, (12, 0), (), "PercentWorkComplete", None),
		"PhoneticType": (89, 2, (3, 0), (), "PhoneticType", None),
		"Priority": (125, 2, (12, 0), (), "Priority", None),
		"ProjectFinish": (8, 2, (12, 0), (), "ProjectFinish", None),
		"ProjectGuideContent": (6029, 2, (8, 0), (), "ProjectGuideContent", None),
		"ProjectGuideFunctionalLayoutPage": (6025, 2, (8, 0), (), "ProjectGuideFunctionalLayoutPage", None),
		"ProjectGuideSaveBuffer": (6028, 2, (8, 0), (), "ProjectGuideSaveBuffer", None),
		"ProjectGuideUseDefaultContent": (6047, 2, (11, 0), (), "ProjectGuideUseDefaultContent", None),
		"ProjectGuideUseDefaultFunctionalLayoutPage": (6046, 2, (11, 0), (), "ProjectGuideUseDefaultFunctionalLayoutPage", None),
		"ProjectNotes": (6, 2, (8, 0), (), "ProjectNotes", None),
		"ProjectServerUsedForTracking": (6035, 2, (11, 0), (), "ProjectServerUsedForTracking", None),
		"ProjectStart": (7, 2, (12, 0), (), "ProjectStart", None),
		"PublishInformationOnSave": (6030, 2, (3, 0), (), "PublishInformationOnSave", None),
		"ReceiveNotifications": (95, 2, (11, 0), (), "ReceiveNotifications", None),
		"RemoveFileProperties": (6054, 2, (11, 0), (), "RemoveFileProperties", None),
		"Rollup": (182, 2, (12, 0), (), "Rollup", None),
		"ScheduleFromStart": (12, 2, (11, 0), (), "ScheduleFromStart", None),
		"SendHyperlinkNote": (96, 2, (11, 0), (), "SendHyperlinkNote", None),
		"ServerIdentification": (6016, 2, (3, 0), (), "ServerIdentification", None),
		"ServerPath": (94, 2, (8, 0), (), "ServerPath", None),
		"ServerURL": (93, 2, (8, 0), (), "ServerURL", None),
		"ShowCriticalSlack": (5107, 2, (3, 0), (), "ShowCriticalSlack", None),
		"ShowCrossProjectLinksInfo": (87, 2, (11, 0), (), "ShowCrossProjectLinksInfo", None),
		"ShowEstimatedDuration": (6004, 2, (11, 0), (), "ShowEstimatedDuration", None),
		"ShowExternalPredecessors": (86, 2, (11, 0), (), "ShowExternalPredecessors", None),
		"ShowExternalSuccessors": (85, 2, (11, 0), (), "ShowExternalSuccessors", None),
		"SpaceBeforeTimeLabels": (61, 2, (11, 0), (), "SpaceBeforeTimeLabels", None),
		"SpreadCostsToStatusDate": (83, 2, (11, 0), (), "SpreadCostsToStatusDate", None),
		"SpreadPercentCompleteToStatusDate": (90, 2, (11, 0), (), "SpreadPercentCompleteToStatusDate", None),
		"Start1": (152, 2, (12, 0), (), "Start1", None),
		"Start2": (155, 2, (12, 0), (), "Start2", None),
		"Start3": (158, 2, (12, 0), (), "Start3", None),
		"Start4": (161, 2, (12, 0), (), "Start4", None),
		"Start5": (164, 2, (12, 0), (), "Start5", None),
		"StartOnCurrentDate": (5110, 2, (11, 0), (), "StartOnCurrentDate", None),
		"StartWeekOn": (5126, 2, (3, 0), (), "StartWeekOn", None),
		"StartYearIn": (5127, 2, (3, 0), (), "StartYearIn", None),
		"StatusDate": (75, 2, (12, 0), (), "StatusDate", None),
		"Subject": (28, 2, (12, 0), (), "Subject", None),
		"TaskErrorCount": (6107, 2, (3, 0), (), "TaskErrorCount", None),
		"TeamMembersCanDeclineTasks": (6002, 2, (11, 0), (), "TeamMembersCanDeclineTasks", None),
		"Text1": (151, 2, (8, 0), (), "Text1", None),
		"Text10": (170, 2, (8, 0), (), "Text10", None),
		"Text2": (154, 2, (8, 0), (), "Text2", None),
		"Text3": (157, 2, (8, 0), (), "Text3", None),
		"Text4": (160, 2, (8, 0), (), "Text4", None),
		"Text5": (163, 2, (8, 0), (), "Text5", None),
		"Text6": (166, 2, (8, 0), (), "Text6", None),
		"Text7": (167, 2, (8, 0), (), "Text7", None),
		"Text8": (168, 2, (8, 0), (), "Text8", None),
		"Text9": (169, 2, (8, 0), (), "Text9", None),
		"Title": (27, 2, (12, 0), (), "Title", None),
		"TrackOvertimeWork": (6001, 2, (11, 0), (), "TrackOvertimeWork", None),
		"TrackingMethod": (6036, 2, (3, 0), (), "TrackingMethod", None),
		"UnderlineHyperlinks": (99, 2, (11, 0), (), "UnderlineHyperlinks", None),
		"UpdateProjOnSave": (6017, 2, (11, 0), (), "UpdateProjOnSave", None),
		"UseFYStartYear": (64, 2, (11, 0), (), "UseFYStartYear", None),
		"VBASigned": (6018, 2, (11, 0), (), "VBASigned", None),
		"WBS": (116, 2, (8, 0), (), "WBS", None),
		"WBSCodeGenerate": (6013, 2, (11, 0), (), "WBSCodeGenerate", None),
		"WBSVerifyUniqueness": (6014, 2, (11, 0), (), "WBSVerifyUniqueness", None),
		"WeekLabelDisplay": (59, 2, (2, 0), (), "WeekLabelDisplay", None),
		# Property 'Windows' is an object of type 'Windows'
		"Windows": (30, 2, (9, 0), (), "Windows", '{00020B03-0000-0000-C000-000000000046}'),
		# Property 'Windows2' is an object of type 'Windows2'
		"Windows2": (6056, 2, (9, 0), (), "Windows2", '{00020B05-0000-0000-C000-000000000046}'),
		"WorkgroupMessages": (92, 2, (3, 0), (), "WorkgroupMessages", None),
		"YearLabelDisplay": (60, 2, (2, 0), (), "YearLabelDisplay", None),
		"ActualCost": (107, 2, (12, 0), (), "ActualCost", None),
		"ActualDuration": (128, 2, (12, 0), (), "ActualDuration", None),
		"ActualFinish": (142, 2, (12, 0), (), "ActualFinish", None),
		"ActualStart": (141, 2, (12, 0), (), "ActualStart", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (65524, 2, (13, 0), (), "Application", '{36D27C48-A1E8-11D3-BA55-00C04F72F325}'),
		"BCWP": (111, 2, (12, 0), (), "BCWP", None),
		"BCWS": (112, 2, (12, 0), (), "BCWS", None),
		# Method 'BaseCalendars' returns object of type 'Calendars'
		"BaseCalendars": (10, 2, (9, 0), (), "BaseCalendars", '{000C0C44-0000-0000-C000-000000000046}'),
		"BaselineDuration": (127, 2, (12, 0), (), "BaselineDuration", None),
		"BuiltinDocumentProperties": (53, 2, (9, 0), (), "BuiltinDocumentProperties", None),
		"CV": (183, 2, (12, 0), (), "CV", None),
		# Method 'Calendar' returns object of type 'Calendar'
		"Calendar": (51, 2, (9, 0), (), "Calendar", '{000C0C43-0000-0000-C000-000000000046}'),
		"CanCheckIn": (6052, 2, (11, 0), (), "CanCheckIn", None),
		"CodeName": (72, 2, (8, 0), (), "CodeName", None),
		# Method 'CommandBars' returns object of type 'CommandBars'
		"CommandBars": (68, 2, (13, 0), (), "CommandBars", '{55F88893-7708-11D1-ACEB-006008961DA5}'),
		"Confirmed": (210, 2, (12, 0), (), "Confirmed", None),
		"ConstraintDate": (118, 2, (12, 0), (), "ConstraintDate", None),
		"ConstraintType": (117, 2, (12, 0), (), "ConstraintType", None),
		"Container": (55, 2, (9, 0), (), "Container", None),
		"Cost": (105, 2, (12, 0), (), "Cost", None),
		"CostVariance": (109, 2, (12, 0), (), "CostVariance", None),
		"Created": (193, 2, (12, 0), (), "Created", None),
		"CreationDate": (22, 2, (12, 0), (), "CreationDate", None),
		"Critical": (119, 2, (12, 0), (), "Critical", None),
		"CurrentFilter": (39, 2, (8, 0), (), "CurrentFilter", None),
		"CurrentGroup": (6008, 2, (8, 0), (), "CurrentGroup", None),
		"CurrentTable": (38, 2, (8, 0), (), "CurrentTable", None),
		"CurrentView": (37, 2, (8, 0), (), "CurrentView", None),
		"CustomDocumentProperties": (54, 2, (9, 0), (), "CustomDocumentProperties", None),
		# Method 'DetectCycle' returns object of type 'Tasks'
		"DetectCycle": (6128, 2, (9, 0), (), "DetectCycle", '{000C0C40-0000-0000-C000-000000000046}'),
		# Method 'DocumentLibraryVersions' returns object of type 'DocumentLibraryVersions'
		"DocumentLibraryVersions": (6050, 2, (9, 0), (), "DocumentLibraryVersions", '{000C0388-0000-0000-C000-000000000046}'),
		"Duration": (129, 2, (12, 0), (), "Duration", None),
		"DurationVariance": (130, 2, (12, 0), (), "DurationVariance", None),
		"EarlyFinish": (138, 2, (12, 0), (), "EarlyFinish", None),
		"EarlyStart": (137, 2, (12, 0), (), "EarlyStart", None),
		"Finish": (136, 2, (12, 0), (), "Finish", None),
		"FinishVariance": (146, 2, (12, 0), (), "FinishVariance", None),
		"FreeSlack": (121, 2, (12, 0), (), "FreeSlack", None),
		"FullName": (42, 2, (8, 0), (), "FullName", None),
		"HasPassword": (5105, 2, (11, 0), (), "HasPassword", None),
		"ID": (47, 2, (3, 0), (), "ID", None),
		"Index": (65527, 2, (12, 0), (), "Index", None),
		"KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled": (6124, 2, (11, 0), (), "KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled", None),
		"LastPrintedDate": (23, 2, (12, 0), (), "LastPrintedDate", None),
		"LastSaveDate": (24, 2, (12, 0), (), "LastSaveDate", None),
		"LastSavedBy": (25, 2, (8, 0), (), "LastSavedBy", None),
		"LateFinish": (140, 2, (12, 0), (), "LateFinish", None),
		"LateStart": (139, 2, (12, 0), (), "LateStart", None),
		"LinkedFields": (198, 2, (12, 0), (), "LinkedFields", None),
		"ManuallyScheduledTasksAutoRespectLinks": (6123, 2, (11, 0), (), "ManuallyScheduledTasksAutoRespectLinks", None),
		# Method 'MapList' returns object of type 'List'
		"MapList": (6006, 2, (9, 0), (), "MapList", '{00020B17-0000-0000-C000-000000000046}'),
		"Milestone": (124, 2, (12, 0), (), "Milestone", None),
		"NewTasksCreatedAsManual": (6117, 2, (11, 0), (), "NewTasksCreatedAsManual", None),
		"NumberOfResources": (41, 2, (3, 0), (), "NumberOfResources", None),
		"NumberOfTasks": (40, 2, (3, 0), (), "NumberOfTasks", None),
		"Objects": (197, 2, (3, 0), (), "Objects", None),
		# Method 'OutlineChildren' returns object of type 'Tasks'
		"OutlineChildren": (52, 2, (9, 0), (), "OutlineChildren", '{000C0C40-0000-0000-C000-000000000046}'),
		# Method 'OutlineCodes' returns object of type 'OutlineCodes'
		"OutlineCodes": (6048, 2, (9, 0), (), "OutlineCodes", '{4269779B-F035-4E2F-ABDD-54B6D94A2A03}'),
		"OutlineLevel": (185, 2, (3, 0), (), "OutlineLevel", None),
		"OutlineNumber": (202, 2, (8, 0), (), "OutlineNumber", None),
		"Parent": (65523, 2, (9, 0), (), "Parent", None),
		"Path": (11, 2, (8, 0), (), "Path", None),
		"PercentComplete": (132, 2, (12, 0), (), "PercentComplete", None),
		"Project": (184, 2, (12, 0), (), "Project", None),
		"ProjectNamePrefix": (6033, 2, (8, 0), (), "ProjectNamePrefix", None),
		# Method 'ProjectSummaryTask' returns object of type 'Task'
		"ProjectSummaryTask": (80, 2, (9, 0), (), "ProjectSummaryTask", '{000C0C3F-0000-0000-C000-000000000046}'),
		"ReadOnly": (5104, 2, (11, 0), (), "ReadOnly", None),
		"ReadOnlyRecommended": (46, 2, (11, 0), (), "ReadOnlyRecommended", None),
		"RemainingCost": (110, 2, (12, 0), (), "RemainingCost", None),
		"RemainingDuration": (131, 2, (12, 0), (), "RemainingDuration", None),
		"RemainingWork": (104, 2, (12, 0), (), "RemainingWork", None),
		# Method 'ReportList' returns object of type 'List'
		"ReportList": (32, 2, (9, 0), (), "ReportList", '{00020B17-0000-0000-C000-000000000046}'),
		# Method 'ResourceFilterList' returns object of type 'List'
		"ResourceFilterList": (34, 2, (9, 0), (), "ResourceFilterList", '{00020B17-0000-0000-C000-000000000046}'),
		# Method 'ResourceFilters' returns object of type 'Filters'
		"ResourceFilters": (6032, 2, (9, 0), (), "ResourceFilters", '{ED989E98-F561-4D58-8F67-5D2E2B920E77}'),
		"ResourceGroup": (213, 2, (12, 0), (), "ResourceGroup", None),
		# Method 'ResourceGroupList' returns object of type 'List'
		"ResourceGroupList": (6010, 2, (9, 0), (), "ResourceGroupList", '{00020B17-0000-0000-C000-000000000046}'),
		# Method 'ResourceGroups' returns object of type 'ResourceGroups'
		"ResourceGroups": (6012, 2, (9, 0), (), "ResourceGroups", '{11589055-69F0-11D2-B889-00C04FB90729}'),
		# Method 'ResourceGroups2' returns object of type 'ResourceGroups2'
		"ResourceGroups2": (6122, 2, (9, 0), (), "ResourceGroups2", '{1158905B-69F0-11D2-B889-00C04FB90729}'),
		"ResourceInitials": (150, 2, (12, 0), (), "ResourceInitials", None),
		"ResourceNames": (149, 2, (12, 0), (), "ResourceNames", None),
		"ResourcePoolName": (43, 2, (8, 0), (), "ResourcePoolName", None),
		# Method 'ResourceTableList' returns object of type 'List'
		"ResourceTableList": (36, 2, (9, 0), (), "ResourceTableList", '{00020B17-0000-0000-C000-000000000046}'),
		# Method 'ResourceTables' returns object of type 'Tables'
		"ResourceTables": (6024, 2, (9, 0), (), "ResourceTables", '{31E3EB5A-6339-43B0-B1B8-1AED03886AEC}'),
		# Method 'ResourceViewList' returns object of type 'List'
		"ResourceViewList": (45, 2, (9, 0), (), "ResourceViewList", '{00020B17-0000-0000-C000-000000000046}'),
		# Method 'Resources' returns object of type 'Resources'
		"Resources": (2, 2, (9, 0), (), "Resources", '{000C0C42-0000-0000-C000-000000000046}'),
		"Resume": (199, 2, (12, 0), (), "Resume", None),
		"ResumeNoEarlierThan": (201, 2, (12, 0), (), "ResumeNoEarlierThan", None),
		"RevisionNumber": (26, 2, (8, 0), (), "RevisionNumber", None),
		"SV": (113, 2, (12, 0), (), "SV", None),
		"Saved": (50, 2, (11, 0), (), "Saved", None),
		# Method 'SharedWorkspace' returns object of type 'SharedWorkspace'
		"SharedWorkspace": (6049, 2, (9, 0), (), "SharedWorkspace", '{000C0385-0000-0000-C000-000000000046}'),
		"ShowTaskSuggestions": (6126, 2, (11, 0), (), "ShowTaskSuggestions", None),
		"ShowTaskWarnings": (6125, 2, (11, 0), (), "ShowTaskWarnings", None),
		"Start": (135, 2, (12, 0), (), "Start", None),
		"StartVariance": (145, 2, (12, 0), (), "StartVariance", None),
		"Stop": (200, 2, (12, 0), (), "Stop", None),
		# Method 'Subprojects' returns object of type 'Subprojects'
		"Subprojects": (6007, 2, (9, 0), (), "Subprojects", '{00020B1B-0000-0000-C000-000000000046}'),
		"Summary": (192, 2, (12, 0), (), "Summary", None),
		# Method 'TaskFilterList' returns object of type 'List'
		"TaskFilterList": (33, 2, (9, 0), (), "TaskFilterList", '{00020B17-0000-0000-C000-000000000046}'),
		# Method 'TaskFilters' returns object of type 'Filters'
		"TaskFilters": (6031, 2, (9, 0), (), "TaskFilters", '{ED989E98-F561-4D58-8F67-5D2E2B920E77}'),
		# Method 'TaskGroupList' returns object of type 'List'
		"TaskGroupList": (6009, 2, (9, 0), (), "TaskGroupList", '{00020B17-0000-0000-C000-000000000046}'),
		# Method 'TaskGroups' returns object of type 'TaskGroups'
		"TaskGroups": (6011, 2, (9, 0), (), "TaskGroups", '{11589054-69F0-11D2-B889-00C04FB90729}'),
		# Method 'TaskGroups2' returns object of type 'TaskGroups2'
		"TaskGroups2": (6121, 2, (9, 0), (), "TaskGroups2", '{1158905A-69F0-11D2-B889-00C04FB90729}'),
		# Method 'TaskTableList' returns object of type 'List'
		"TaskTableList": (35, 2, (9, 0), (), "TaskTableList", '{00020B17-0000-0000-C000-000000000046}'),
		# Method 'TaskTables' returns object of type 'Tables'
		"TaskTables": (6023, 2, (9, 0), (), "TaskTables", '{31E3EB5A-6339-43B0-B1B8-1AED03886AEC}'),
		# Method 'TaskViewList' returns object of type 'List'
		"TaskViewList": (44, 2, (9, 0), (), "TaskViewList", '{00020B17-0000-0000-C000-000000000046}'),
		# Method 'Tasks' returns object of type 'Tasks'
		"Tasks": (1, 2, (9, 0), (), "Tasks", '{000C0C40-0000-0000-C000-000000000046}'),
		"TempToDoList": (6038, 2, (3, 0), (), "TempToDoList", None),
		"Template": (29, 2, (8, 0), (), "Template", None),
		"TotalSlack": (122, 2, (12, 0), (), "TotalSlack", None),
		"Type": (6039, 2, (3, 0), (), "Type", None),
		"UniqueID": (186, 2, (3, 0), (), "UniqueID", None),
		"UpdateNeeded": (211, 2, (12, 0), (), "UpdateNeeded", None),
		"UserControl": (70, 2, (11, 0), (), "UserControl", None),
		# Method 'VBProject' returns object of type 'VBProject'
		"VBProject": (69, 2, (13, 0), (), "VBProject", '{0002E169-0000-0000-C000-000000000046}'),
		"VersionName": (6034, 2, (8, 0), (), "VersionName", None),
		# Method 'ViewList' returns object of type 'List'
		"ViewList": (31, 2, (9, 0), (), "ViewList", '{00020B17-0000-0000-C000-000000000046}'),
		# Method 'Views' returns object of type 'Views'
		"Views": (6021, 2, (9, 0), (), "Views", '{4CF32482-106B-4FFF-A1AB-7DD395FB0958}'),
		# Method 'ViewsCombination' returns object of type 'ViewsCombination'
		"ViewsCombination": (6027, 2, (9, 0), (), "ViewsCombination", '{CE4F7D83-369B-43CF-96A8-29C2DE2B8658}'),
		# Method 'ViewsSingle' returns object of type 'ViewsSingle'
		"ViewsSingle": (6026, 2, (9, 0), (), "ViewsSingle", '{B4097887-EC12-4683-9622-9974EF26353C}'),
		"Work": (100, 2, (12, 0), (), "Work", None),
		"WorkVariance": (103, 2, (12, 0), (), "WorkVariance", None),
		"WriteReserved": (5106, 2, (11, 0), (), "WriteReserved", None),
		"_CodeName": (-2147418112, 2, (8, 0), (), "_CodeName", None),
	}
	_prop_map_put_ = {
		"AcceptNewExternalData" : ((88, LCID, 4, 0),()),
		"ActualWork" : ((102, LCID, 4, 0),()),
		"AdministrativeProject" : ((6055, LCID, 4, 0),()),
		"AllowTaskDelegation" : ((5129, LCID, 4, 0),()),
		"AndMoveCompleted" : ((6043, LCID, 4, 0),()),
		"AndMoveRemaining" : ((6041, LCID, 4, 0),()),
		"AskForCompletedWork" : ((6000, LCID, 4, 0),()),
		"Author" : ((48, LCID, 4, 0),()),
		"AutoAddResources" : ((5122, LCID, 4, 0),()),
		"AutoCalcCosts" : ((84, LCID, 4, 0),()),
		"AutoFilter" : ((66, LCID, 4, 0),()),
		"AutoLinkTasks" : ((5114, LCID, 4, 0),()),
		"AutoSplitTasks" : ((5113, LCID, 4, 0),()),
		"AutoTrack" : ((5112, LCID, 4, 0),()),
		"BaselineCost" : ((106, LCID, 4, 0),()),
		"BaselineFinish" : ((144, LCID, 4, 0),()),
		"BaselineStart" : ((143, LCID, 4, 0),()),
		"BaselineWork" : ((101, LCID, 4, 0),()),
		"Comments" : ((14, LCID, 4, 0),()),
		"Company" : ((4, LCID, 4, 0),()),
		"Contact" : ((212, LCID, 4, 0),()),
		"Cost1" : ((206, LCID, 4, 0),()),
		"Cost2" : ((207, LCID, 4, 0),()),
		"Cost3" : ((208, LCID, 4, 0),()),
		"CurrencyCode" : ((5131, LCID, 4, 0),()),
		"CurrencyDigits" : ((5103, LCID, 4, 0),()),
		"CurrencySymbol" : ((5101, LCID, 4, 0),()),
		"CurrencySymbolPosition" : ((5102, LCID, 4, 0),()),
		"CurrentDate" : ((9, LCID, 4, 0),()),
		"DatabaseProjectUniqueID" : ((6019, LCID, 4, 0),()),
		"DayLabelDisplay" : ((58, LCID, 4, 0),()),
		"DaysPerMonth" : ((5128, LCID, 4, 0),()),
		"DefaultDurationUnits" : ((5108, LCID, 4, 0),()),
		"DefaultEarnedValueMethod" : ((6044, LCID, 4, 0),()),
		"DefaultEffortDriven" : ((63, LCID, 4, 0),()),
		"DefaultFinishTime" : ((5123, LCID, 4, 0),()),
		"DefaultFixedCostAccrual" : ((81, LCID, 4, 0),()),
		"DefaultResourceOvertimeRate" : ((5120, LCID, 4, 0),()),
		"DefaultResourceStandardRate" : ((5119, LCID, 4, 0),()),
		"DefaultStartTime" : ((5115, LCID, 4, 0),()),
		"DefaultTaskType" : ((67, LCID, 4, 0),()),
		"DefaultWorkUnits" : ((5109, LCID, 4, 0),()),
		"Delay" : ((120, LCID, 4, 0),()),
		"DisplayProjectSummaryTask" : ((5121, LCID, 4, 0),()),
		"Duration1" : ((203, LCID, 4, 0),()),
		"Duration2" : ((204, LCID, 4, 0),()),
		"Duration3" : ((205, LCID, 4, 0),()),
		"EarnedValueBaseline" : ((6045, LCID, 4, 0),()),
		"EnterpriseActualsSynched" : ((6051, LCID, 4, 0),()),
		"ExpandDatabaseTimephasedData" : ((6020, LCID, 4, 0),()),
		"Finish1" : ((153, LCID, 4, 0),()),
		"Finish2" : ((156, LCID, 4, 0),()),
		"Finish3" : ((159, LCID, 4, 0),()),
		"Finish4" : ((162, LCID, 4, 0),()),
		"Finish5" : ((165, LCID, 4, 0),()),
		"FixedCost" : ((108, LCID, 4, 0),()),
		"FixedDuration" : ((134, LCID, 4, 0),()),
		"Flag1" : ((172, LCID, 4, 0),()),
		"Flag10" : ((181, LCID, 4, 0),()),
		"Flag2" : ((173, LCID, 4, 0),()),
		"Flag3" : ((174, LCID, 4, 0),()),
		"Flag4" : ((175, LCID, 4, 0),()),
		"Flag5" : ((176, LCID, 4, 0),()),
		"Flag6" : ((177, LCID, 4, 0),()),
		"Flag7" : ((178, LCID, 4, 0),()),
		"Flag8" : ((179, LCID, 4, 0),()),
		"Flag9" : ((180, LCID, 4, 0),()),
		"FollowedHyperlinkColor" : ((98, LCID, 4, 0),()),
		"FollowedHyperlinkColorEx" : ((6120, LCID, 4, 0),()),
		"HideBar" : ((209, LCID, 4, 0),()),
		"HonorConstraints" : ((73, LCID, 4, 0),()),
		"HourLabelDisplay" : ((57, LCID, 4, 0),()),
		"HoursPerDay" : ((5117, LCID, 4, 0),()),
		"HoursPerWeek" : ((5118, LCID, 4, 0),()),
		"HyperlinkColor" : ((97, LCID, 4, 0),()),
		"HyperlinkColorEx" : ((6119, LCID, 4, 0),()),
		"IsTemplate" : ((6111, LCID, 4, 0),()),
		"Keywords" : ((49, LCID, 4, 0),()),
		"LevelEntireProject" : ((91, LCID, 4, 0),()),
		"LevelFromDate" : ((76, LCID, 4, 0),()),
		"LevelToDate" : ((77, LCID, 4, 0),()),
		"Manager" : ((5, LCID, 4, 0),()),
		"Marked" : ((171, LCID, 4, 0),()),
		"MinuteLabelDisplay" : ((56, LCID, 4, 0),()),
		"MonthLabelDisplay" : ((6015, LCID, 4, 0),()),
		"MoveCompleted" : ((6040, LCID, 4, 0),()),
		"MoveRemaining" : ((6042, LCID, 4, 0),()),
		"MultipleCriticalPaths" : ((74, LCID, 4, 0),()),
		"Name" : ((0, LCID, 4, 0),()),
		"NewTasksEstimated" : ((6005, LCID, 4, 0),()),
		"Notes" : ((115, LCID, 4, 0),()),
		"Number1" : ((187, LCID, 4, 0),()),
		"Number2" : ((188, LCID, 4, 0),()),
		"Number3" : ((189, LCID, 4, 0),()),
		"Number4" : ((190, LCID, 4, 0),()),
		"Number5" : ((191, LCID, 4, 0),()),
		"PercentWorkComplete" : ((133, LCID, 4, 0),()),
		"PhoneticType" : ((89, LCID, 4, 0),()),
		"Priority" : ((125, LCID, 4, 0),()),
		"ProjectFinish" : ((8, LCID, 4, 0),()),
		"ProjectGuideContent" : ((6029, LCID, 4, 0),()),
		"ProjectGuideFunctionalLayoutPage" : ((6025, LCID, 4, 0),()),
		"ProjectGuideSaveBuffer" : ((6028, LCID, 4, 0),()),
		"ProjectGuideUseDefaultContent" : ((6047, LCID, 4, 0),()),
		"ProjectGuideUseDefaultFunctionalLayoutPage" : ((6046, LCID, 4, 0),()),
		"ProjectNotes" : ((6, LCID, 4, 0),()),
		"ProjectServerUsedForTracking" : ((6035, LCID, 4, 0),()),
		"ProjectStart" : ((7, LCID, 4, 0),()),
		"PublishInformationOnSave" : ((6030, LCID, 4, 0),()),
		"ReceiveNotifications" : ((95, LCID, 4, 0),()),
		"RemoveFileProperties" : ((6054, LCID, 4, 0),()),
		"Rollup" : ((182, LCID, 4, 0),()),
		"ScheduleFromStart" : ((12, LCID, 4, 0),()),
		"SendHyperlinkNote" : ((96, LCID, 4, 0),()),
		"ServerIdentification" : ((6016, LCID, 4, 0),()),
		"ServerPath" : ((94, LCID, 4, 0),()),
		"ServerURL" : ((93, LCID, 4, 0),()),
		"ShowCriticalSlack" : ((5107, LCID, 4, 0),()),
		"ShowCrossProjectLinksInfo" : ((87, LCID, 4, 0),()),
		"ShowEstimatedDuration" : ((6004, LCID, 4, 0),()),
		"ShowExternalPredecessors" : ((86, LCID, 4, 0),()),
		"ShowExternalSuccessors" : ((85, LCID, 4, 0),()),
		"SpaceBeforeTimeLabels" : ((61, LCID, 4, 0),()),
		"SpreadCostsToStatusDate" : ((83, LCID, 4, 0),()),
		"SpreadPercentCompleteToStatusDate" : ((90, LCID, 4, 0),()),
		"Start1" : ((152, LCID, 4, 0),()),
		"Start2" : ((155, LCID, 4, 0),()),
		"Start3" : ((158, LCID, 4, 0),()),
		"Start4" : ((161, LCID, 4, 0),()),
		"Start5" : ((164, LCID, 4, 0),()),
		"StartOnCurrentDate" : ((5110, LCID, 4, 0),()),
		"StartWeekOn" : ((5126, LCID, 4, 0),()),
		"StartYearIn" : ((5127, LCID, 4, 0),()),
		"StatusDate" : ((75, LCID, 4, 0),()),
		"Subject" : ((28, LCID, 4, 0),()),
		"TaskErrorCount" : ((6107, LCID, 4, 0),()),
		"TeamMembersCanDeclineTasks" : ((6002, LCID, 4, 0),()),
		"Text1" : ((151, LCID, 4, 0),()),
		"Text10" : ((170, LCID, 4, 0),()),
		"Text2" : ((154, LCID, 4, 0),()),
		"Text3" : ((157, LCID, 4, 0),()),
		"Text4" : ((160, LCID, 4, 0),()),
		"Text5" : ((163, LCID, 4, 0),()),
		"Text6" : ((166, LCID, 4, 0),()),
		"Text7" : ((167, LCID, 4, 0),()),
		"Text8" : ((168, LCID, 4, 0),()),
		"Text9" : ((169, LCID, 4, 0),()),
		"Title" : ((27, LCID, 4, 0),()),
		"TrackOvertimeWork" : ((6001, LCID, 4, 0),()),
		"TrackingMethod" : ((6036, LCID, 4, 0),()),
		"UnderlineHyperlinks" : ((99, LCID, 4, 0),()),
		"UpdateProjOnSave" : ((6017, LCID, 4, 0),()),
		"UseFYStartYear" : ((64, LCID, 4, 0),()),
		"VBASigned" : ((6018, LCID, 4, 0),()),
		"WBS" : ((116, LCID, 4, 0),()),
		"WBSCodeGenerate" : ((6013, LCID, 4, 0),()),
		"WBSVerifyUniqueness" : ((6014, LCID, 4, 0),()),
		"WeekLabelDisplay" : ((59, LCID, 4, 0),()),
		"Windows" : ((30, LCID, 4, 0),()),
		"Windows2" : ((6056, LCID, 4, 0),()),
		"WorkgroupMessages" : ((92, LCID, 4, 0),()),
		"YearLabelDisplay" : ((60, LCID, 4, 0),()),
		"KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled": ((6124, LCID, 4, 0),()),
		"ManuallyScheduledTasksAutoRespectLinks": ((6123, LCID, 4, 0),()),
		"NewTasksCreatedAsManual": ((6117, LCID, 4, 0),()),
		"ShowTaskSuggestions": ((6126, LCID, 4, 0),()),
		"ShowTaskWarnings": ((6125, LCID, 4, 0),()),
		"TempToDoList": ((6038, LCID, 4, 0),()),
		"_CodeName": ((-2147418112, LCID, 4, 0),()),
	}
	# Default property for this class is 'Name'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (8, 0), (), "Name", None))
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
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{00020B00-0000-0000-C000-000000000046}", _IProjectDoc )
