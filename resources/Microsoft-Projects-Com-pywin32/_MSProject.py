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

from win32com.client import DispatchBaseClass
class _MSProject(DispatchBaseClass):
	CLSID = IID('{00020AFF-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{36D27C48-A1E8-11D3-BA55-00C04F72F325}')

	def About(self):
		return self._oleobj_.InvokeTypes(2376, LCID, 1, (11, 0), (),)

	def ActivateMicrosoftApp(self, Index=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5052, LCID, 1, (24, 0), ((3, 0),),Index
			)

	def AddNewColumn(self, Column=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(710, LCID, 1, (11, 0), ((12, 16),),Column
			)

	def AddProgressLine(self):
		return self._oleobj_.InvokeTypes(252, LCID, 1, (11, 0), (),)

	def AddResourcesFromProjectServer(self):
		return self._oleobj_.InvokeTypes(2130, LCID, 1, (11, 0), (),)

	def AfterUnloadWebBrowserControl(self):
		return self._oleobj_.InvokeTypes(5232, LCID, 1, (24, 0), (),)

	def Alerts(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(10, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def AppExecute(self, Window=defaultNamedOptArg, Command=defaultNamedOptArg, Minimize=defaultNamedOptArg, Activate=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(8, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Window
			, Command, Minimize, Activate)

	def AppLaunch(self, Application=defaultNamedNotOptArg, Document=defaultNamedOptArg, Activate=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(20, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16)),Application
			, Document, Activate)

	def AppMaximize(self):
		return self._oleobj_.InvokeTypes(2008, LCID, 1, (11, 0), (),)

	def AppMinimize(self):
		return self._oleobj_.InvokeTypes(2009, LCID, 1, (11, 0), (),)

	def AppMove(self, XPosition=defaultNamedOptArg, YPosition=defaultNamedOptArg, Points=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2010, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),XPosition
			, YPosition, Points)

	def AppRestore(self):
		return self._oleobj_.InvokeTypes(2011, LCID, 1, (11, 0), (),)

	def AppSize(self, Width=defaultNamedOptArg, Height=defaultNamedOptArg, Points=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2012, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),Width
			, Height, Points)

	def AutoCorrect(self):
		return self._oleobj_.InvokeTypes(627, LCID, 1, (11, 0), (),)

	def AutoFilter(self):
		return self._oleobj_.InvokeTypes(22, LCID, 1, (11, 0), (),)

	def AutoSaveToGlobal(self, OnOff=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1500, LCID, 1, (11, 0), ((12, 16),),OnOff
			)

	def BarBoxFormat(self):
		return self._oleobj_.InvokeTypes(2389, LCID, 1, (11, 0), (),)

	def BarBoxStyles(self):
		return self._oleobj_.InvokeTypes(904, LCID, 1, (11, 0), (),)

	def BarRounding(self, On=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2080, LCID, 1, (11, 0), ((12, 16),),On
			)

	def BaseCalendarCreate(self, Name=defaultNamedNotOptArg, FromName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(618, LCID, 1, (11, 0), ((8, 0), (12, 16)),Name
			, FromName)

	def BaseCalendarDelete(self, Name=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(619, LCID, 1, (11, 0), ((8, 0),),Name
			)

	def BaseCalendarEditDays(self, Name=defaultNamedNotOptArg, StartDate=defaultNamedOptArg, EndDate=defaultNamedOptArg, WeekDay=defaultNamedOptArg
			, Working=defaultNamedOptArg, From1=defaultNamedOptArg, To1=defaultNamedOptArg, From2=defaultNamedOptArg, To2=defaultNamedOptArg
			, From3=defaultNamedOptArg, To3=defaultNamedOptArg, Default=defaultNamedOptArg, From4=defaultNamedOptArg, To4=defaultNamedOptArg
			, From5=defaultNamedOptArg, To5=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(615, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, StartDate, EndDate, WeekDay, Working, From1
			, To1, From2, To2, From3, To3
			, Default, From4, To4, From5, To5
			)

	def BaseCalendarRename(self, FromName=defaultNamedNotOptArg, ToName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(624, LCID, 1, (11, 0), ((8, 0), (8, 0)),FromName
			, ToName)

	def BaseCalendarReset(self, Name=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(617, LCID, 1, (11, 0), ((8, 0),),Name
			)

	def BaseCalendars(self, Index=defaultNamedOptArg, Locked=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(604, LCID, 1, (11, 0), ((12, 16), (12, 16)),Index
			, Locked)

	def BaselineClear(self, All=defaultNamedOptArg, From=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2384, LCID, 1, (11, 0), ((12, 16), (12, 16)),All
			, From)

	def BaselineSave(self, All=defaultNamedOptArg, Copy=defaultNamedOptArg, Into=defaultNamedOptArg, RollupToSummaryTasks=defaultNamedOptArg
			, RollupFromSubtasks=defaultNamedOptArg, SetDefaults=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(610, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),All
			, Copy, Into, RollupToSummaryTasks, RollupFromSubtasks, SetDefaults
			)

	def BoxAlign(self, Alignment=-1):
		return self._oleobj_.InvokeTypes(29, LCID, 1, (11, 0), ((3, 48),),Alignment
			)

	def BoxCellEdit(self, Name=defaultNamedNotOptArg, Cell=defaultNamedNotOptArg, FieldName=-1, Font=defaultNamedNotOptArg
			, FontSize=defaultNamedNotOptArg, FontColor=-1, Bold=defaultNamedNotOptArg, Italic=defaultNamedNotOptArg, Underline=defaultNamedNotOptArg
			, HorizontalAlignment=-1, VerticalAlignment=-1, TextLineLimit=defaultNamedNotOptArg, ShowLabel=defaultNamedNotOptArg, Label=defaultNamedNotOptArg
			, DateFormat=-1):
		return self._oleobj_.InvokeTypes(2393, LCID, 1, (11, 0), ((8, 0), (3, 0), (3, 48), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (3, 48), (3, 48), (12, 16), (12, 16), (12, 16), (3, 48)),Name
			, Cell, FieldName, Font, FontSize, FontColor
			, Bold, Italic, Underline, HorizontalAlignment, VerticalAlignment
			, TextLineLimit, ShowLabel, Label, DateFormat)

	def BoxCellEditEx(self, Name=defaultNamedNotOptArg, Cell=defaultNamedNotOptArg, FieldName=-1, Font=defaultNamedNotOptArg
			, FontSize=defaultNamedNotOptArg, FontColor=defaultNamedNotOptArg, Bold=defaultNamedNotOptArg, Italic=defaultNamedNotOptArg, Underline=defaultNamedNotOptArg
			, HorizontalAlignment=-1, VerticalAlignment=-1, TextLineLimit=defaultNamedNotOptArg, ShowLabel=defaultNamedNotOptArg, Label=defaultNamedNotOptArg
			, DateFormat=-1):
		return self._oleobj_.InvokeTypes(2156, LCID, 1, (11, 0), ((8, 0), (3, 0), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (3, 48), (3, 48), (12, 16), (12, 16), (12, 16), (3, 48)),Name
			, Cell, FieldName, Font, FontSize, FontColor
			, Bold, Italic, Underline, HorizontalAlignment, VerticalAlignment
			, TextLineLimit, ShowLabel, Label, DateFormat)

	def BoxCellLayout(self, Name=defaultNamedNotOptArg, CellRows=defaultNamedOptArg, CellColumns=defaultNamedOptArg, CellWidth=defaultNamedOptArg
			, MergeCells=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2392, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, CellRows, CellColumns, CellWidth, MergeCells)

	def BoxDataTemplate(self, Name=defaultNamedNotOptArg, action=defaultNamedNotOptArg, NewName=defaultNamedOptArg, Overwrite=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2391, LCID, 1, (11, 0), ((8, 0), (3, 0), (12, 16), (12, 16)),Name
			, action, NewName, Overwrite)

	def BoxFormat(self, ProjectName=defaultNamedNotOptArg, TaskID=defaultNamedNotOptArg, DataTemplate=defaultNamedNotOptArg, HorizontalGridlines=defaultNamedNotOptArg
			, VerticalGridlines=defaultNamedNotOptArg, BorderShape=-1, BorderColor=-1, BorderWidth=defaultNamedNotOptArg, BackgroundColor=-1
			, BackgroundPattern=-1, Reset=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2388, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (3, 48), (3, 48), (12, 16), (3, 48), (3, 48), (12, 16)),ProjectName
			, TaskID, DataTemplate, HorizontalGridlines, VerticalGridlines, BorderShape
			, BorderColor, BorderWidth, BackgroundColor, BackgroundPattern, Reset
			)

	def BoxFormatEx(self, ProjectName=defaultNamedNotOptArg, TaskID=defaultNamedNotOptArg, DataTemplate=defaultNamedNotOptArg, HorizontalGridlines=defaultNamedNotOptArg
			, VerticalGridlines=defaultNamedNotOptArg, BorderShape=-1, BorderColor=defaultNamedNotOptArg, BorderWidth=defaultNamedNotOptArg, BackgroundColor=defaultNamedNotOptArg
			, BackgroundPattern=-1, Reset=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2155, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (3, 48), (12, 16)),ProjectName
			, TaskID, DataTemplate, HorizontalGridlines, VerticalGridlines, BorderShape
			, BorderColor, BorderWidth, BackgroundColor, BackgroundPattern, Reset
			)

	def BoxGetXPosition(self, TaskID=defaultNamedNotOptArg, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(5057, LCID, 1, (3, 0), ((3, 0), (12, 16)),TaskID
			, ProjectName)

	def BoxGetYPosition(self, TaskID=defaultNamedNotOptArg, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(5058, LCID, 1, (3, 0), ((3, 0), (12, 16)),TaskID
			, ProjectName)

	def BoxLayout(self, LayoutMode=-1, LayoutScheme=-1, SummaryPrecedence=defaultNamedNotOptArg, RowAlignment=-1
			, ColumnAlignment=-1, RowSpacing=defaultNamedNotOptArg, ColumnSpacing=defaultNamedNotOptArg, RowHeight=-1, ColumnWidth=-1
			, AdjustForPageBreaks=defaultNamedNotOptArg, ShowSummaryTasks=defaultNamedNotOptArg, ViewBackgroundColor=-1, ViewBackgroundPattern=-1, ShowProgressMarks=defaultNamedOptArg
			, ShowPageBreaks=defaultNamedOptArg, ShowIDOnly=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(43, LCID, 1, (11, 0), ((3, 48), (3, 48), (12, 16), (3, 48), (3, 48), (12, 16), (12, 16), (3, 48), (3, 48), (12, 16), (12, 16), (3, 48), (3, 48), (12, 16), (12, 16), (12, 16)),LayoutMode
			, LayoutScheme, SummaryPrecedence, RowAlignment, ColumnAlignment, RowSpacing
			, ColumnSpacing, RowHeight, ColumnWidth, AdjustForPageBreaks, ShowSummaryTasks
			, ViewBackgroundColor, ViewBackgroundPattern, ShowProgressMarks, ShowPageBreaks, ShowIDOnly
			)

	def BoxLayoutEx(self, LayoutMode=-1, LayoutScheme=-1, SummaryPrecedence=defaultNamedNotOptArg, RowAlignment=-1
			, ColumnAlignment=-1, RowSpacing=defaultNamedNotOptArg, ColumnSpacing=defaultNamedNotOptArg, RowHeight=-1, ColumnWidth=-1
			, AdjustForPageBreaks=defaultNamedNotOptArg, ShowSummaryTasks=defaultNamedNotOptArg, ViewBackgroundColor=defaultNamedNotOptArg, ViewBackgroundPattern=-1, ShowProgressMarks=defaultNamedOptArg
			, ShowPageBreaks=defaultNamedOptArg, ShowIDOnly=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2157, LCID, 1, (11, 0), ((3, 48), (3, 48), (12, 16), (3, 48), (3, 48), (12, 16), (12, 16), (3, 48), (3, 48), (12, 16), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16)),LayoutMode
			, LayoutScheme, SummaryPrecedence, RowAlignment, ColumnAlignment, RowSpacing
			, ColumnSpacing, RowHeight, ColumnWidth, AdjustForPageBreaks, ShowSummaryTasks
			, ViewBackgroundColor, ViewBackgroundPattern, ShowProgressMarks, ShowPageBreaks, ShowIDOnly
			)

	def BoxLinkLabelsShow(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(47, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def BoxLinkStyleToggle(self, StraightLinks=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(48, LCID, 1, (11, 0), ((12, 16),),StraightLinks
			)

	def BoxLinks(self, Style=-1, ShowArrows=defaultNamedNotOptArg, ShowLabels=defaultNamedNotOptArg, ColorMode=-1
			, CriticalColor=1, NoncriticalColor=0):
		return self._oleobj_.InvokeTypes(44, LCID, 1, (11, 0), ((3, 48), (12, 16), (12, 16), (3, 48), (3, 48), (3, 48)),Style
			, ShowArrows, ShowLabels, ColorMode, CriticalColor, NoncriticalColor
			)

	def BoxLinksEx(self, Style=-1, ShowArrows=defaultNamedNotOptArg, ShowLabels=defaultNamedNotOptArg, ColorMode=-1
			, CriticalColor=defaultNamedOptArg, NoncriticalColor=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2158, LCID, 1, (11, 0), ((3, 48), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16)),Style
			, ShowArrows, ShowLabels, ColorMode, CriticalColor, NoncriticalColor
			)

	def BoxProgressMarksShow(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(46, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def BoxSet(self, action=4, TaskID=defaultNamedOptArg, XPosition=defaultNamedOptArg, YPosition=defaultNamedOptArg
			, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(49, LCID, 1, (11, 0), ((3, 48), (12, 16), (12, 16), (12, 16), (12, 16)),action
			, TaskID, XPosition, YPosition, ProjectName)

	def BoxShowHideFields(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(905, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def BoxStylesEdit(self, Style=defaultNamedNotOptArg, DataTemplate=defaultNamedNotOptArg, HorizontalGridlines=defaultNamedNotOptArg, VerticalGridlines=defaultNamedNotOptArg
			, BorderShape=-1, BorderColor=-1, BorderWidth=defaultNamedNotOptArg, BackgroundColor=-1, BackgroundPattern=-1):
		return self._oleobj_.InvokeTypes(2387, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16), (3, 48), (3, 48), (12, 16), (3, 48), (3, 48)),Style
			, DataTemplate, HorizontalGridlines, VerticalGridlines, BorderShape, BorderColor
			, BorderWidth, BackgroundColor, BackgroundPattern)

	def BoxStylesEditEx(self, Style=defaultNamedNotOptArg, DataTemplate=defaultNamedNotOptArg, HorizontalGridlines=defaultNamedNotOptArg, VerticalGridlines=defaultNamedNotOptArg
			, BorderShape=-1, BorderColor=defaultNamedNotOptArg, BorderWidth=defaultNamedNotOptArg, BackgroundColor=defaultNamedNotOptArg, BackgroundPattern=-1):
		return self._oleobj_.InvokeTypes(2154, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (3, 48)),Style
			, DataTemplate, HorizontalGridlines, VerticalGridlines, BorderShape, BorderColor
			, BorderWidth, BackgroundColor, BackgroundPattern)

	def BoxZoom(self, Percent=defaultNamedOptArg, Entire=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(308, LCID, 1, (11, 0), ((12, 16), (12, 16)),Percent
			, Entire)

	def CacheSettings(self):
		return self._oleobj_.InvokeTypes(2276, LCID, 1, (11, 0), (),)

	def CacheStatus(self):
		return self._oleobj_.InvokeTypes(2282, LCID, 1, (11, 0), (),)

	def CalculateAll(self):
		return self._oleobj_.InvokeTypes(607, LCID, 1, (11, 0), (),)

	def CalculateProject(self):
		return self._oleobj_.InvokeTypes(2034, LCID, 1, (11, 0), (),)

	def CalendarBarStyles(self, BarRounding=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2326, LCID, 1, (11, 0), ((12, 16),),BarRounding
			)

	def CalendarBarStylesEdit(self, Item=defaultNamedNotOptArg, Bar=defaultNamedOptArg, Pattern=defaultNamedOptArg, Color=defaultNamedOptArg
			, Align=defaultNamedOptArg, Wrap=defaultNamedOptArg, Shadow=defaultNamedOptArg, Field1=defaultNamedOptArg, Field2=defaultNamedOptArg
			, Field3=defaultNamedOptArg, Field4=defaultNamedOptArg, Field5=defaultNamedOptArg, SplitPattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2339, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, Bar, Pattern, Color, Align, Wrap
			, Shadow, Field1, Field2, Field3, Field4
			, Field5, SplitPattern)

	def CalendarBarStylesEditEx(self, Item=defaultNamedNotOptArg, Bar=defaultNamedOptArg, Pattern=defaultNamedOptArg, Color=defaultNamedOptArg
			, Align=defaultNamedOptArg, Wrap=defaultNamedOptArg, Shadow=defaultNamedOptArg, Field1=defaultNamedOptArg, Field2=defaultNamedOptArg
			, Field3=defaultNamedOptArg, Field4=defaultNamedOptArg, Field5=defaultNamedOptArg, SplitPattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2146, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, Bar, Pattern, Color, Align, Wrap
			, Shadow, Field1, Field2, Field3, Field4
			, Field5, SplitPattern)

	def CalendarBestFitWeekHeight(self):
		return self._oleobj_.InvokeTypes(2327, LCID, 1, (11, 0), (),)

	def CalendarDateBoxes(self, TopLeft=defaultNamedOptArg, TopRight=defaultNamedOptArg, BottomLeft=defaultNamedOptArg, BottomRight=defaultNamedOptArg
			, TopColor=defaultNamedOptArg, BottomColor=defaultNamedOptArg, TopPattern=defaultNamedOptArg, BottomPattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2340, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),TopLeft
			, TopRight, BottomLeft, BottomRight, TopColor, BottomColor
			, TopPattern, BottomPattern)

	def CalendarDateBoxesEx(self, TopLeft=defaultNamedOptArg, TopRight=defaultNamedOptArg, BottomLeft=defaultNamedOptArg, BottomRight=defaultNamedOptArg
			, TopColor=defaultNamedOptArg, BottomColor=defaultNamedOptArg, TopPattern=defaultNamedOptArg, BottomPattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2148, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),TopLeft
			, TopRight, BottomLeft, BottomRight, TopColor, BottomColor
			, TopPattern, BottomPattern)

	def CalendarDateShading(self, BaseCalendarName=defaultNamedOptArg, ResourceUniqueID=defaultNamedOptArg, ProjectIndex=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2344, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),BaseCalendarName
			, ResourceUniqueID, ProjectIndex)

	def CalendarDateShadingEdit(self, Item=defaultNamedNotOptArg, Pattern=defaultNamedOptArg, Color=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2343, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16)),Item
			, Pattern, Color)

	def CalendarDateShadingEditEx(self, Item=defaultNamedNotOptArg, Pattern=defaultNamedOptArg, Color=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2147, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16)),Item
			, Pattern, Color)

	def CalendarLayout(self, SortOrder=defaultNamedOptArg, AutoLayout=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2346, LCID, 1, (11, 0), ((12, 16), (12, 16)),SortOrder
			, AutoLayout)

	def CalendarShowBarSplits(self, Display=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2382, LCID, 1, (11, 0), ((12, 16),),Display
			)

	def CalendarTaskList(self, Date=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2354, LCID, 1, (11, 0), ((12, 16),),Date
			)

	def CalendarTimescale(self):
		return self._oleobj_.InvokeTypes(2345, LCID, 1, (11, 0), (),)

	def CalendarWeekHeadings(self, MonthTitle=defaultNamedOptArg, WeekTitle=defaultNamedOptArg, DayTitle=defaultNamedOptArg, ShowPreview=defaultNamedOptArg
			, DaysPerWeek=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2341, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),MonthTitle
			, WeekTitle, DayTitle, ShowPreview, DaysPerWeek)

	def CalendarWeekHeadingsEx(self, MonthTitle=defaultNamedOptArg, WeekTitle=defaultNamedOptArg, DayTitle=defaultNamedOptArg, ShowPreview=defaultNamedOptArg
			, DaysPerWeek=defaultNamedOptArg, ShowTitleBeginningEndDates=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2269, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),MonthTitle
			, WeekTitle, DayTitle, ShowPreview, DaysPerWeek, ShowTitleBeginningEndDates
			)

	def ChangeColumnDataType(self, Type=0, Column=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(711, LCID, 1, (11, 0), ((3, 48), (12, 16)),Type
			, Column)

	def ChangeStatusDate(self, Date=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2181, LCID, 1, (11, 0), ((12, 16),),Date
			)

	def ChangeWorkingTime(self, CalendarName=defaultNamedOptArg, Locked=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(625, LCID, 1, (11, 0), ((12, 16), (12, 16)),CalendarName
			, Locked)

	def ChangeWorkingTimeEx(self, CalendarName=defaultNamedOptArg, Locked=defaultNamedOptArg, SelectedDate=defaultNamedOptArg, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2270, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),CalendarName
			, Locked, SelectedDate, ProjectName)

	def CheckField(self, Field=defaultNamedNotOptArg, Value=defaultNamedNotOptArg, Test=defaultNamedOptArg, Op=defaultNamedOptArg
			, Field2=defaultNamedOptArg, Value2=defaultNamedOptArg, Test2=defaultNamedOptArg):
		return self._ApplyTypes_(7, 1, (12, 0), ((8, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)), u'CheckField', None,Field
			, Value, Test, Op, Field2, Value2
			, Test2)

	def CheckIn(self, fSaveChanges=defaultNamedOptArg, Comments=defaultNamedOptArg, fMakePublic=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2323, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),fSaveChanges
			, Comments, fMakePublic)

	def CheckOut(self):
		return self._oleobj_.InvokeTypes(2332, LCID, 1, (11, 0), (),)

	def CheckResourceErrors(self, LocalRUID=defaultNamedOptArg, ResetImport=defaultNamedOptArg, CheckEnterprise=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2258, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),LocalRUID
			, ResetImport, CheckEnterprise)

	def CheckTaskErrors(self, TaskID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2257, LCID, 1, (11, 0), ((12, 16),),TaskID
			)

	def CleanupCache(self):
		return self._oleobj_.InvokeTypes(2277, LCID, 1, (11, 0), (),)

	def CleanupProjectFromCache(self, Filename=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2191, LCID, 1, (11, 0), ((12, 16),),Filename
			)

	def ClearConstraint(self, TaskID=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1514, LCID, 1, (11, 0), ((3, 0),),TaskID
			)

	def ClipboardShow(self):
		return self._oleobj_.InvokeTypes(707, LCID, 1, (11, 0), (),)

	def CloseComparison(self):
		return self._oleobj_.InvokeTypes(2187, LCID, 1, (11, 0), (),)

	def CloseUndoTransaction(self):
		return self._oleobj_.InvokeTypes(5220, LCID, 1, (24, 0), (),)

	def ColumnAlignment(self, Align=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2325, LCID, 1, (11, 0), ((3, 0),),Align
			)

	def ColumnBestFit(self, Column=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2037, LCID, 1, (11, 0), ((12, 16),),Column
			)

	def ColumnDelete(self):
		return self._oleobj_.InvokeTypes(230, LCID, 1, (11, 0), (),)

	def ColumnEdit(self, Column=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2038, LCID, 1, (11, 0), ((12, 16),),Column
			)

	def ColumnInsert(self):
		return self._oleobj_.InvokeTypes(228, LCID, 1, (11, 0), (),)

	def ComAddInsDialog(self):
		return self._oleobj_.InvokeTypes(2395, LCID, 1, (11, 0), (),)

	def CommitmentsPane(self):
		return self._oleobj_.InvokeTypes(2280, LCID, 1, (11, 0), (),)

	def CompareProjectVersions(self):
		return self._oleobj_.InvokeTypes(2183, LCID, 1, (11, 0), (),)

	def CompareProjectsLegendToggle(self):
		return self._oleobj_.InvokeTypes(2188, LCID, 1, (11, 0), (),)

	def ConsolidateProjects(self, Filenames=defaultNamedNotOptArg, NewWindow=defaultNamedNotOptArg, AttachToSources=defaultNamedNotOptArg, PoolResources=defaultNamedNotOptArg
			, HideSubtasks=defaultNamedNotOptArg, openPool=0, UserID=defaultNamedOptArg, Password=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(124, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16)),Filenames
			, NewWindow, AttachToSources, PoolResources, HideSubtasks, openPool
			, UserID, Password)

	def ConvertHangulToHanja(self):
		return self._oleobj_.InvokeTypes(28, LCID, 1, (11, 0), (),)

	def CreateComparisonReport(self, Filename=defaultNamedNotOptArg, TaskTable=defaultNamedNotOptArg, ResourceTable=defaultNamedNotOptArg, Items=6
			, Columns=0, ShowLegend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2182, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (3, 48), (3, 48), (12, 16)),Filename
			, TaskTable, ResourceTable, Items, Columns, ShowLegend
			)

	def CreateEnterpriseCalendar(self):
		return self._oleobj_.InvokeTypes(2135, LCID, 1, (11, 0), (),)

	def CreatePublisher(self, Edition=defaultNamedOptArg, Contains=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(222, LCID, 1, (11, 0), ((12, 16), (12, 16)),Edition
			, Contains)

	def CreateWebAccount(self, ServerURL=defaultNamedNotOptArg, Name=defaultNamedNotOptArg, AuthenticationType=0, AccountType=0
			, ShowDialog=defaultNamedOptArg, Email=defaultNamedOptArg, WindowsAccount=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2390, LCID, 1, (11, 0), ((12, 16), (12, 16), (3, 48), (3, 48), (12, 16), (12, 16), (12, 16)),ServerURL
			, Name, AuthenticationType, AccountType, ShowDialog, Email
			, WindowsAccount)

	def CustomFieldDelete(self, Field=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2132, LCID, 1, (11, 0), ((3, 0),),Field
			)

	def CustomFieldGetFormula(self, FieldID=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5059, LCID, 1, (8, 0), ((3, 0),),FieldID
			)

	def CustomFieldGetName(self, FieldID=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5039, LCID, 1, (8, 0), ((3, 0),),FieldID
			)

	def CustomFieldIndicatorAdd(self, FieldID=defaultNamedNotOptArg, Test=defaultNamedNotOptArg, Value=defaultNamedNotOptArg, IndicatorID=defaultNamedNotOptArg
			, CriteriaList=0, Index=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(38, LCID, 1, (11, 0), ((3, 0), (3, 0), (8, 0), (3, 0), (3, 48), (12, 16)),FieldID
			, Test, Value, IndicatorID, CriteriaList, Index
			)

	def CustomFieldIndicatorDelete(self, FieldID=defaultNamedNotOptArg, Index=defaultNamedNotOptArg, CriteriaList=0):
		return self._oleobj_.InvokeTypes(39, LCID, 1, (11, 0), ((3, 0), (2, 0), (3, 48)),FieldID
			, Index, CriteriaList)

	def CustomFieldIndicators(self, FieldID=defaultNamedNotOptArg, SummaryInheritsNonsummary=defaultNamedOptArg, ProjectInheritsSummary=defaultNamedOptArg, ShowToolTips=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(37, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16)),FieldID
			, SummaryInheritsNonsummary, ProjectInheritsSummary, ShowToolTips)

	def CustomFieldMappingDialog(self, TaskCustomFields=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2256, LCID, 1, (11, 0), ((12, 16),),TaskCustomFields
			)

	def CustomFieldProperties(self, FieldID=defaultNamedNotOptArg, Attribute=-1, SummaryCalc=-1, GraphicalIndicators=defaultNamedOptArg
			, Required=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(35, LCID, 1, (11, 0), ((3, 0), (3, 48), (3, 48), (12, 16), (12, 16)),FieldID
			, Attribute, SummaryCalc, GraphicalIndicators, Required)

	def CustomFieldPropertiesEx(self, FieldID=defaultNamedNotOptArg, Attribute=-1, SummaryCalc=-1, GraphicalIndicators=defaultNamedOptArg
			, Required=defaultNamedOptArg, AutomaticallyRolldownToAssn=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2267, LCID, 1, (11, 0), ((3, 0), (3, 48), (3, 48), (12, 16), (12, 16), (12, 16)),FieldID
			, Attribute, SummaryCalc, GraphicalIndicators, Required, AutomaticallyRolldownToAssn
			)

	def CustomFieldRename(self, FieldID=defaultNamedNotOptArg, NewName=defaultNamedOptArg, Phonetic=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2378, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16)),FieldID
			, NewName, Phonetic)

	def CustomFieldSetFormula(self, FieldID=defaultNamedNotOptArg, Formula=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(36, LCID, 1, (11, 0), ((3, 0), (12, 16)),FieldID
			, Formula)

	def CustomFieldValueList(self, FieldID=defaultNamedNotOptArg, ListDefault=defaultNamedNotOptArg, DefaultValue=defaultNamedNotOptArg, RestrictToList=defaultNamedNotOptArg
			, AppendNew=defaultNamedNotOptArg, PromptOnNew=defaultNamedNotOptArg, DisplayOrder=-1):
		return self._oleobj_.InvokeTypes(40, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (3, 48)),FieldID
			, ListDefault, DefaultValue, RestrictToList, AppendNew, PromptOnNew
			, DisplayOrder)

	def CustomFieldValueListAdd(self, FieldID=defaultNamedNotOptArg, Value=defaultNamedOptArg, Description=defaultNamedOptArg, Phonetic=defaultNamedOptArg
			, Index=defaultNamedOptArg, FieldDefault=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(41, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),FieldID
			, Value, Description, Phonetic, Index, FieldDefault
			)

	def CustomFieldValueListDelete(self, FieldID=defaultNamedNotOptArg, Index=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(42, LCID, 1, (11, 0), ((3, 0), (2, 0)),FieldID
			, Index)

	def CustomFieldValueListGetItem(self, FieldID=defaultNamedNotOptArg, Item=defaultNamedNotOptArg, Index=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5060, LCID, 1, (8, 0), ((3, 0), (3, 0), (3, 0)),FieldID
			, Item, Index)

	def CustomForms(self):
		return self._oleobj_.InvokeTypes(1003, LCID, 1, (11, 0), (),)

	def CustomOutlineCodeEdit(self, FieldID=defaultNamedNotOptArg, Level=defaultNamedNotOptArg, Sequence=-1, Length=defaultNamedOptArg
			, Separator=defaultNamedOptArg, OnlyLookUpTableCodes=defaultNamedOptArg, OnlyCompleteCodes=defaultNamedOptArg, LookupTableLink=defaultNamedOptArg, OnlyLeaves=defaultNamedOptArg
			, MatchGeneric=defaultNamedOptArg, RequiredCode=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(631, LCID, 1, (11, 0), ((3, 0), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),FieldID
			, Level, Sequence, Length, Separator, OnlyLookUpTableCodes
			, OnlyCompleteCodes, LookupTableLink, OnlyLeaves, MatchGeneric, RequiredCode
			)

	def CustomOutlineCodeEditEx(self, FieldID=defaultNamedNotOptArg, Level=defaultNamedNotOptArg, Sequence=-1, Length=defaultNamedOptArg
			, Separator=defaultNamedOptArg, OnlyLookUpTableCodes=defaultNamedOptArg, OnlyCompleteCodes=defaultNamedOptArg, LookupTableLink=defaultNamedOptArg, OnlyLeaves=defaultNamedOptArg
			, MatchGeneric=defaultNamedOptArg, RequiredCode=defaultNamedOptArg, LookupDefault=defaultNamedOptArg, DefaultValue=defaultNamedOptArg, SortOrder=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2265, LCID, 1, (11, 0), ((3, 0), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),FieldID
			, Level, Sequence, Length, Separator, OnlyLookUpTableCodes
			, OnlyCompleteCodes, LookupTableLink, OnlyLeaves, MatchGeneric, RequiredCode
			, LookupDefault, DefaultValue, SortOrder)

	def CustomizeField(self):
		return self._oleobj_.InvokeTypes(2379, LCID, 1, (11, 0), (),)

	def CustomizeIMEMode(self, FieldID=-1, IMEMode=-1):
		return self._oleobj_.InvokeTypes(254, LCID, 1, (11, 0), ((3, 48), (3, 48)),FieldID
			, IMEMode)

	def DDEExecute(self, Command=defaultNamedNotOptArg, TimeOut=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1202, LCID, 1, (11, 0), ((8, 0), (12, 16)),Command
			, TimeOut)

	def DDEInitiate(self, App=defaultNamedNotOptArg, Topic=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1201, LCID, 1, (11, 0), ((8, 0), (8, 0)),App
			, Topic)

	def DDELinksUpdate(self):
		return self._oleobj_.InvokeTypes(2053, LCID, 1, (11, 0), (),)

	def DDEPasteLink(self):
		return self._oleobj_.InvokeTypes(207, LCID, 1, (11, 0), (),)

	def DDETerminate(self):
		return self._oleobj_.InvokeTypes(1203, LCID, 1, (11, 0), (),)

	def DateAdd(self, StartDate=defaultNamedNotOptArg, Duration=defaultNamedNotOptArg, Calendar=defaultNamedOptArg):
		return self._ApplyTypes_(5023, 1, (12, 0), ((12, 0), (12, 0), (12, 16)), u'DateAdd', None,StartDate
			, Duration, Calendar)

	def DateDifference(self, StartDate=defaultNamedNotOptArg, FinishDate=defaultNamedNotOptArg, Calendar=defaultNamedOptArg):
		return self._ApplyTypes_(5025, 1, (12, 0), ((12, 0), (12, 0), (12, 16)), u'DateDifference', None,StartDate
			, FinishDate, Calendar)

	def DateFormat(self, Date=defaultNamedNotOptArg, Format=defaultNamedOptArg):
		return self._ApplyTypes_(5031, 1, (12, 0), ((12, 0), (12, 16)), u'DateFormat', None,Date
			, Format)

	def DateSubtract(self, FinishDate=defaultNamedNotOptArg, Duration=defaultNamedNotOptArg, Calendar=defaultNamedOptArg):
		return self._ApplyTypes_(5024, 1, (12, 0), ((12, 0), (12, 0), (12, 16)), u'DateSubtract', None,FinishDate
			, Duration, Calendar)

	def DeleteFromDatabase(self, Name=defaultNamedOptArg, UserID=defaultNamedOptArg, DatabasePassWord=defaultNamedOptArg, FormatID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(135, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Name
			, UserID, DatabasePassWord, FormatID)

	def DependenciesPane(self):
		return self._oleobj_.InvokeTypes(2281, LCID, 1, (11, 0), (),)

	def DetailStylesAdd(self, Item=0, Position=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(963, LCID, 1, (11, 0), ((3, 48), (12, 16)),Item
			, Position)

	def DetailStylesFormat(self, Item=defaultNamedOptArg, Font=defaultNamedOptArg, Size=defaultNamedOptArg, Bold=defaultNamedOptArg
			, Italic=defaultNamedOptArg, Underline=defaultNamedOptArg, Color=defaultNamedOptArg, CellColor=defaultNamedOptArg, Pattern=defaultNamedOptArg
			, ShowInMenu=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(962, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, Font, Size, Bold, Italic, Underline
			, Color, CellColor, Pattern, ShowInMenu)

	def DetailStylesFormatEx(self, Item=defaultNamedOptArg, Font=defaultNamedOptArg, Size=defaultNamedOptArg, Bold=defaultNamedOptArg
			, Italic=defaultNamedOptArg, Underline=defaultNamedOptArg, Color=defaultNamedOptArg, CellColor=defaultNamedOptArg, Pattern=defaultNamedOptArg
			, ShowInMenu=defaultNamedOptArg, Strikethrough=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2164, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, Font, Size, Bold, Italic, Underline
			, Color, CellColor, Pattern, ShowInMenu, Strikethrough
			)

	def DetailStylesProperties(self, AlignCellData=defaultNamedOptArg, RepeatRowLabel=defaultNamedOptArg, ShortLabels=defaultNamedOptArg, DisplayDetailsColumn=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(952, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),AlignCellData
			, RepeatRowLabel, ShortLabels, DisplayDetailsColumn)

	def DetailStylesRemove(self, Item=0):
		return self._oleobj_.InvokeTypes(964, LCID, 1, (11, 0), ((3, 48),),Item
			)

	def DetailStylesRemoveAll(self):
		return self._oleobj_.InvokeTypes(965, LCID, 1, (11, 0), (),)

	def DetailStylesToggleItem(self, Item=0):
		return self._oleobj_.InvokeTypes(960, LCID, 1, (11, 0), ((3, 48),),Item
			)

	def DetailsPaneToggle(self, Timeline=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(85, LCID, 1, (11, 0), ((12, 16),),Timeline
			)

	def DisplaySharedWorkspace(self):
		return self._oleobj_.InvokeTypes(2363, LCID, 1, (11, 0), (),)

	def DocClose(self):
		return self._oleobj_.InvokeTypes(2007, LCID, 1, (11, 0), (),)

	def DocMaximize(self):
		return self._oleobj_.InvokeTypes(2013, LCID, 1, (11, 0), (),)

	def DocMove(self, XPosition=defaultNamedOptArg, YPosition=defaultNamedOptArg, Points=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2015, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),XPosition
			, YPosition, Points)

	def DocRestore(self):
		return self._oleobj_.InvokeTypes(2016, LCID, 1, (11, 0), (),)

	def DocSize(self, Width=defaultNamedOptArg, Height=defaultNamedOptArg, Points=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2017, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),Width
			, Height, Points)

	def DocumentExport(self, Filename=defaultNamedNotOptArg, FileType=0, IncludeDocumentProperties=defaultNamedOptArg, IncludeDocumentMarkup=defaultNamedOptArg
			, ArchiveFormat=defaultNamedOptArg, FromDate=defaultNamedOptArg, ToDate=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2173, LCID, 1, (11, 0), ((12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Filename
			, FileType, IncludeDocumentProperties, IncludeDocumentMarkup, ArchiveFormat, FromDate
			, ToDate)

	def DocumentLibraryVersionsDialog(self):
		return self._oleobj_.InvokeTypes(2342, LCID, 1, (11, 0), (),)

	def DrawingCreate(self, Type=defaultNamedNotOptArg, Behind=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2306, LCID, 1, (11, 0), ((3, 0), (12, 16)),Type
			, Behind)

	def DrawingCycleColor(self):
		return self._oleobj_.InvokeTypes(2315, LCID, 1, (11, 0), (),)

	def DrawingMove(self, Forward=defaultNamedOptArg, Full=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2311, LCID, 1, (11, 0), ((12, 16), (12, 16)),Forward
			, Full)

	def DrawingProperties(self, SizePositionTab=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2307, LCID, 1, (11, 0), ((12, 16),),SizePositionTab
			)

	def DrawingReshape(self):
		return self._oleobj_.InvokeTypes(2314, LCID, 1, (11, 0), (),)

	def DrawingToolbarShow(self):
		return self._oleobj_.InvokeTypes(2352, LCID, 1, (11, 0), (),)

	def DurationFormat(self, Duration=defaultNamedNotOptArg, Units=defaultNamedOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5026, LCID, 1, (8, 0), ((12, 0), (12, 16)),Duration
			, Units)

	def DurationValue(self, Duration=defaultNamedNotOptArg):
		return self._ApplyTypes_(5027, 1, (12, 0), ((8, 0),), u'DurationValue', None,Duration
			)

	def EditClear(self, Contents=defaultNamedOptArg, Formats=defaultNamedOptArg, Notes=defaultNamedOptArg, Hyperlinks=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(205, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Contents
			, Formats, Notes, Hyperlinks)

	def EditClearFormats(self):
		return self._oleobj_.InvokeTypes(240, LCID, 1, (11, 0), (),)

	def EditClearHyperlink(self):
		return self._oleobj_.InvokeTypes(1316, LCID, 1, (11, 0), (),)

	def EditCopy(self):
		return self._oleobj_.InvokeTypes(203, LCID, 1, (11, 0), (),)

	def EditCopyPicture(self, Object=defaultNamedNotOptArg, ForPrinter=defaultNamedNotOptArg, SelectedRows=defaultNamedNotOptArg, FromDate=defaultNamedNotOptArg
			, ToDate=defaultNamedNotOptArg, Filename=defaultNamedNotOptArg, ScaleOption=1, MaxImageHeight=defaultNamedOptArg, MaxImageWidth=defaultNamedOptArg
			, MeasurementUnits=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(204, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16)),Object
			, ForPrinter, SelectedRows, FromDate, ToDate, Filename
			, ScaleOption, MaxImageHeight, MaxImageWidth, MeasurementUnits)

	def EditCut(self):
		return self._oleobj_.InvokeTypes(202, LCID, 1, (11, 0), (),)

	def EditDelete(self):
		return self._oleobj_.InvokeTypes(208, LCID, 1, (11, 0), (),)

	def EditEnterpriseCalendar(self, UniqueID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2134, LCID, 1, (11, 0), ((12, 16),),UniqueID
			)

	def EditGoTo(self, ID=defaultNamedOptArg, Date=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(213, LCID, 1, (11, 0), ((12, 16), (12, 16)),ID
			, Date)

	def EditHyperlink(self, Name=defaultNamedOptArg, Address=defaultNamedOptArg, SubAddress=defaultNamedOptArg, ScreenTip=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1310, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Address, SubAddress, ScreenTip)

	def EditInsert(self):
		return self._oleobj_.InvokeTypes(209, LCID, 1, (11, 0), (),)

	def EditPaste(self):
		return self._oleobj_.InvokeTypes(206, LCID, 1, (11, 0), (),)

	def EditPasteAsHyperlink(self):
		return self._oleobj_.InvokeTypes(1308, LCID, 1, (11, 0), (),)

	def EditPasteSpecial(self, Link=defaultNamedOptArg, Type=defaultNamedOptArg, DisplayAsIcon=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(232, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),Link
			, Type, DisplayAsIcon)

	def EditRedo(self):
		return self._oleobj_.InvokeTypes(200, LCID, 1, (11, 0), (),)

	def EditTPStyle(self, Style=defaultNamedNotOptArg, FillColor=defaultNamedOptArg, BorderColor=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(57, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16)),Style
			, FillColor, BorderColor)

	def EditUndo(self):
		return self._oleobj_.InvokeTypes(201, LCID, 1, (11, 0), (),)

	def EditionStopAll(self, Stop=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(224, LCID, 1, (11, 0), ((12, 16),),Stop
			)

	def EnterpriseCustomOutlineCodeShare(self, LinkFrom=defaultNamedNotOptArg, LinkTo=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2336, LCID, 1, (11, 0), ((3, 0), (12, 16)),LinkFrom
			, LinkTo)

	def EnterpriseCustomizeFields(self):
		return self._oleobj_.InvokeTypes(2125, LCID, 1, (11, 0), (),)

	def EnterpriseGlobalBackup(self, BackupFileName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2250, LCID, 1, (11, 0), ((12, 16),),BackupFileName
			)

	def EnterpriseGlobalCheckOut(self):
		return self._oleobj_.InvokeTypes(2337, LCID, 1, (11, 0), (),)

	def EnterpriseGlobalRestore(self, ProfileName=defaultNamedOptArg, RestoreFileName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2251, LCID, 1, (11, 0), ((12, 16), (12, 16)),ProfileName
			, RestoreFileName)

	def EnterpriseMakeServerURLTrusted(self):
		return self._oleobj_.InvokeTypes(5085, LCID, 1, (24, 0), (),)

	def EnterpriseProjectDelete(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2128, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def EnterpriseProjectImportWizard(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2248, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def EnterpriseProjectProfiles(self):
		return self._oleobj_.InvokeTypes(2317, LCID, 1, (11, 0), (),)

	def EnterpriseResSubstitutionWizard(self, ProjectList=defaultNamedNotOptArg, PoolOption=0, RBSorResourceList=defaultNamedOptArg, FreezeHorizonDate=defaultNamedOptArg
			, UpdateProjects=defaultNamedOptArg, SaveReport=defaultNamedOptArg, Path=defaultNamedOptArg, AssignProposedResources=defaultNamedOptArg, LevelProposedBookings=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2123, LCID, 1, (11, 0), ((12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),ProjectList
			, PoolOption, RBSorResourceList, FreezeHorizonDate, UpdateProjects, SaveReport
			, Path, AssignProposedResources, LevelProposedBookings)

	def EnterpriseResourceGet(self, EUID=defaultNamedOptArg, RUID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2094, LCID, 1, (11, 0), ((12, 16), (12, 16)),EUID
			, RUID)

	def EnterpriseResourcesImport(self, LocalRUIDs=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2090, LCID, 1, (11, 0), ((12, 16),),LocalRUIDs
			)

	def EnterpriseResourcesImportEx(self, LocalRUIDs=defaultNamedOptArg, UseImportColumn=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2274, LCID, 1, (11, 0), ((12, 16), (12, 16)),LocalRUIDs
			, UseImportColumn)

	def EnterpriseResourcesOpen(self, EUID=defaultNamedNotOptArg, OpenType=1):
		return self._oleobj_.InvokeTypes(2088, LCID, 1, (11, 0), ((12, 16), (3, 48)),EUID
			, OpenType)

	def EnterpriseSynchActuals(self, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2254, LCID, 1, (11, 0), ((12, 16),),ProjectName
			)

	def EnterpriseTeamBuilder(self):
		return self._oleobj_.InvokeTypes(2124, LCID, 1, (11, 0), (),)

	def FieldConstantToFieldName(self, Field=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5067, LCID, 1, (8, 0), ((3, 0),),Field
			)

	def FieldNameToFieldConstant(self, FieldName=defaultNamedNotOptArg, FieldType=0):
		return self._oleobj_.InvokeTypes(5066, LCID, 1, (3, 0), ((8, 0), (3, 48)),FieldName
			, FieldType)

	# The method FileBuildID is actually a property, but must be used as a method to correctly pass the arguments
	def FileBuildID(self, Name=defaultNamedNotOptArg, UserID=defaultNamedOptArg, DatabasePassWord=defaultNamedOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5054, LCID, 2, (8, 0), ((8, 0), (12, 16), (12, 16)),Name
			, UserID, DatabasePassWord)

	def FileClose(self, Save=2, NoAuto=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(103, LCID, 1, (11, 0), ((3, 48), (12, 16)),Save
			, NoAuto)

	def FileCloseAll(self, Save=2):
		return self._oleobj_.InvokeTypes(104, LCID, 1, (11, 0), ((3, 48),),Save
			)

	def FileCloseAllEx(self, Save=2, CheckIn=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2260, LCID, 1, (11, 0), ((3, 48), (12, 16)),Save
			, CheckIn)

	def FileCloseEx(self, Save=2, NoAuto=defaultNamedOptArg, CheckIn=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2259, LCID, 1, (11, 0), ((3, 48), (12, 16), (12, 16)),Save
			, NoAuto, CheckIn)

	def FileExit(self, Save=2):
		return self._oleobj_.InvokeTypes(114, LCID, 1, (11, 0), ((3, 48),),Save
			)

	# The method FileFormatID is actually a property, but must be used as a method to correctly pass the arguments
	def FileFormatID(self, Name=defaultNamedNotOptArg, UserID=defaultNamedOptArg, DatabasePassWord=defaultNamedOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5053, LCID, 2, (8, 0), ((8, 0), (12, 16), (12, 16)),Name
			, UserID, DatabasePassWord)

	def FileLoadLast(self, Number=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(117, LCID, 1, (11, 0), ((12, 16),),Number
			)

	def FileNew(self, SummaryInfo=defaultNamedOptArg, Template=defaultNamedOptArg, FileNewDialog=defaultNamedOptArg, FileNewWorkpane=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(101, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),SummaryInfo
			, Template, FileNewDialog, FileNewWorkpane)

	def FileOpen(self, Name=defaultNamedNotOptArg, ReadOnly=defaultNamedNotOptArg, Merge=defaultNamedNotOptArg, TaskInformation=defaultNamedNotOptArg
			, Table=defaultNamedNotOptArg, Sheet=defaultNamedNotOptArg, NoAuto=defaultNamedNotOptArg, UserID=defaultNamedNotOptArg, DatabasePassWord=defaultNamedNotOptArg
			, FormatID=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, openPool=0, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg
			, IgnoreReadOnlyRecommended=defaultNamedOptArg, XMLName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(102, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, ReadOnly, Merge, TaskInformation, Table, Sheet
			, NoAuto, UserID, DatabasePassWord, FormatID, Map
			, openPool, Password, WriteResPassword, IgnoreReadOnlyRecommended, XMLName
			)

	def FileOpenEx(self, Name=defaultNamedNotOptArg, ReadOnly=defaultNamedNotOptArg, Merge=defaultNamedNotOptArg, TaskInformation=defaultNamedNotOptArg
			, Table=defaultNamedNotOptArg, Sheet=defaultNamedNotOptArg, NoAuto=defaultNamedNotOptArg, UserID=defaultNamedNotOptArg, DatabasePassWord=defaultNamedNotOptArg
			, FormatID=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, openPool=0, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg
			, IgnoreReadOnlyRecommended=defaultNamedOptArg, XMLName=defaultNamedOptArg, DoNotLoadFromEnterprise=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2273, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, ReadOnly, Merge, TaskInformation, Table, Sheet
			, NoAuto, UserID, DatabasePassWord, FormatID, Map
			, openPool, Password, WriteResPassword, IgnoreReadOnlyRecommended, XMLName
			, DoNotLoadFromEnterprise)

	def FilePageSetup(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(116, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def FilePageSetupCalendar(self, Name=defaultNamedOptArg, MonthsPerPage=defaultNamedOptArg, WeeksPerPage=defaultNamedOptArg, ScreenWeekHeight=defaultNamedOptArg
			, OnlyDaysInMonth=defaultNamedOptArg, OnlyWeeksInMonth=defaultNamedOptArg, MonthPreviews=defaultNamedOptArg, MonthTitle=defaultNamedOptArg, AdditionalTasks=defaultNamedOptArg
			, GroupAdditionalTasks=defaultNamedOptArg, PrintNotes=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2361, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, MonthsPerPage, WeeksPerPage, ScreenWeekHeight, OnlyDaysInMonth, OnlyWeeksInMonth
			, MonthPreviews, MonthTitle, AdditionalTasks, GroupAdditionalTasks, PrintNotes
			)

	def FilePageSetupCalendarText(self, Name=defaultNamedOptArg, Item=defaultNamedOptArg, Font=defaultNamedOptArg, Size=defaultNamedOptArg
			, Bold=defaultNamedOptArg, Italic=defaultNamedOptArg, Underline=defaultNamedOptArg, Color=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2371, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Item, Font, Size, Bold, Italic
			, Underline, Color)

	def FilePageSetupCalendarTextEx(self, Name=defaultNamedOptArg, Item=defaultNamedOptArg, Font=defaultNamedOptArg, Size=defaultNamedOptArg
			, Bold=defaultNamedOptArg, Italic=defaultNamedOptArg, Underline=defaultNamedOptArg, Color=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2162, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Item, Font, Size, Bold, Italic
			, Underline, Color)

	def FilePageSetupFooter(self, Name=defaultNamedNotOptArg, Alignment=1, Text=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2358, LCID, 1, (11, 0), ((12, 16), (3, 48), (12, 16)),Name
			, Alignment, Text)

	def FilePageSetupHeader(self, Name=defaultNamedNotOptArg, Alignment=1, Text=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2357, LCID, 1, (11, 0), ((12, 16), (3, 48), (12, 16)),Name
			, Alignment, Text)

	def FilePageSetupLegend(self, Name=defaultNamedNotOptArg, TextWidth=defaultNamedNotOptArg, LegendOn=defaultNamedNotOptArg, Alignment=1
			, Text=defaultNamedNotOptArg, LabelFontName=defaultNamedNotOptArg, LabelFontSize=defaultNamedNotOptArg, LabelFontBold=defaultNamedNotOptArg, LabelFontItalic=defaultNamedNotOptArg
			, LabelFontUnderline=defaultNamedNotOptArg, LabelFontColor=-1):
		return self._oleobj_.InvokeTypes(2359, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (3, 48)),Name
			, TextWidth, LegendOn, Alignment, Text, LabelFontName
			, LabelFontSize, LabelFontBold, LabelFontItalic, LabelFontUnderline, LabelFontColor
			)

	def FilePageSetupLegendEx(self, Name=defaultNamedNotOptArg, TextWidth=defaultNamedNotOptArg, LegendOn=defaultNamedNotOptArg, Alignment=1
			, Text=defaultNamedOptArg, LabelFontName=defaultNamedOptArg, LabelFontSize=defaultNamedOptArg, LabelFontBold=defaultNamedOptArg, LabelFontItalic=defaultNamedOptArg
			, LabelFontUnderline=defaultNamedOptArg, LabelFontColor=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2161, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, TextWidth, LegendOn, Alignment, Text, LabelFontName
			, LabelFontSize, LabelFontBold, LabelFontItalic, LabelFontUnderline, LabelFontColor
			)

	def FilePageSetupMargins(self, Name=defaultNamedOptArg, Top=defaultNamedOptArg, Bottom=defaultNamedOptArg, Left=defaultNamedOptArg
			, Right=defaultNamedOptArg, Borders=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2356, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Top, Bottom, Left, Right, Borders
			)

	def FilePageSetupPage(self, Name=defaultNamedNotOptArg, Portrait=defaultNamedNotOptArg, PercentScale=defaultNamedNotOptArg, PagesTall=defaultNamedNotOptArg
			, PagesWide=defaultNamedNotOptArg, PaperSize=0, FirstPageNumber=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2355, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (3, 48), (12, 16)),Name
			, Portrait, PercentScale, PagesTall, PagesWide, PaperSize
			, FirstPageNumber)

	def FilePageSetupView(self, Name=defaultNamedOptArg, AllSheetColumns=defaultNamedOptArg, RepeatColumns=defaultNamedOptArg, PrintNotes=defaultNamedOptArg
			, PrintBlankPages=defaultNamedOptArg, BestPageFitTimescale=defaultNamedOptArg, PrintColumnTotals=defaultNamedOptArg, PrintRowTotals=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2360, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, AllSheetColumns, RepeatColumns, PrintNotes, PrintBlankPages, BestPageFitTimescale
			, PrintColumnTotals, PrintRowTotals)

	def FilePrint(self, FromPage=defaultNamedOptArg, ToPage=defaultNamedOptArg, PageBreaks=defaultNamedOptArg, Draft=defaultNamedOptArg
			, Copies=defaultNamedOptArg, FromDate=defaultNamedOptArg, ToDate=defaultNamedOptArg, OnePageWide=defaultNamedOptArg, Preview=defaultNamedOptArg
			, Color=defaultNamedOptArg, ShowIEPrintDialog=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(109, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),FromPage
			, ToPage, PageBreaks, Draft, Copies, FromDate
			, ToDate, OnePageWide, Preview, Color, ShowIEPrintDialog
			)

	def FilePrintPreview(self):
		return self._oleobj_.InvokeTypes(111, LCID, 1, (11, 0), (),)

	def FilePrintSetup(self, Printer=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(113, LCID, 1, (11, 0), ((12, 16),),Printer
			)

	def FileProperties(self):
		return self._oleobj_.InvokeTypes(626, LCID, 1, (11, 0), (),)

	def FileQuit(self, Save=2):
		return self._oleobj_.InvokeTypes(2375, LCID, 1, (11, 0), ((3, 48),),Save
			)

	def FileSave(self):
		return self._oleobj_.InvokeTypes(106, LCID, 1, (11, 0), (),)

	def FileSaveAs(self, Name=defaultNamedNotOptArg, Format=0, Backup=defaultNamedOptArg, ReadOnly=defaultNamedOptArg
			, TaskInformation=defaultNamedOptArg, Filtered=defaultNamedOptArg, Table=defaultNamedOptArg, UserID=defaultNamedOptArg, DatabasePassWord=defaultNamedOptArg
			, FormatID=defaultNamedOptArg, Map=defaultNamedOptArg, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg, ClearBaseline=defaultNamedOptArg
			, ClearActuals=defaultNamedOptArg, ClearResourceRates=defaultNamedOptArg, ClearFixedCosts=defaultNamedOptArg, XMLName=defaultNamedOptArg, ClearConfirmed=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(107, LCID, 1, (11, 0), ((12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Format, Backup, ReadOnly, TaskInformation, Filtered
			, Table, UserID, DatabasePassWord, FormatID, Map
			, Password, WriteResPassword, ClearBaseline, ClearActuals, ClearResourceRates
			, ClearFixedCosts, XMLName, ClearConfirmed)

	def FileSaveOffline(self, Save=2):
		return self._oleobj_.InvokeTypes(136, LCID, 1, (11, 0), ((3, 48),),Save
			)

	def FileSaveWorkspace(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(108, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def FillAcross(self, Right=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(244, LCID, 1, (11, 0), ((12, 16),),Right
			)

	def FillDown(self, Down=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(218, LCID, 1, (11, 0), ((12, 16),),Down
			)

	def FilterApply(self, Name=defaultNamedOptArg, Highlight=defaultNamedOptArg, Value1=defaultNamedOptArg, Value2=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(502, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Highlight, Value1, Value2)

	def FilterClear(self):
		return self._oleobj_.InvokeTypes(505, LCID, 1, (11, 0), (),)

	def FilterEdit(self, Name=defaultNamedNotOptArg, TaskFilter=defaultNamedNotOptArg, Create=defaultNamedOptArg, OverwriteExisting=defaultNamedOptArg
			, Parenthesis=defaultNamedOptArg, NewName=defaultNamedOptArg, FieldName=defaultNamedOptArg, NewFieldName=defaultNamedOptArg, Test=defaultNamedOptArg
			, Value=defaultNamedOptArg, Operation=defaultNamedOptArg, ShowInMenu=defaultNamedOptArg, ShowSummaryTasks=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(503, LCID, 1, (11, 0), ((8, 0), (11, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, TaskFilter, Create, OverwriteExisting, Parenthesis, NewName
			, FieldName, NewFieldName, Test, Value, Operation
			, ShowInMenu, ShowSummaryTasks)

	def FilterNew(self, FilterType=65535):
		return self._oleobj_.InvokeTypes(504, LCID, 1, (11, 0), ((3, 48),),FilterType
			)

	def FilterShowSummaryRows(self, On=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2297, LCID, 1, (11, 0), ((11, 0),),On
			)

	def Filters(self):
		return self._oleobj_.InvokeTypes(501, LCID, 1, (11, 0), (),)

	def Find(self, Field=defaultNamedOptArg, Test=defaultNamedOptArg, Value=defaultNamedOptArg, Next=defaultNamedOptArg
			, MatchCase=defaultNamedOptArg, FieldID=defaultNamedOptArg, TestID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(215, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Field
			, Test, Value, Next, MatchCase, FieldID
			, TestID)

	def FindEx(self, Field=defaultNamedOptArg, Test=defaultNamedOptArg, Value=defaultNamedOptArg, Next=defaultNamedOptArg
			, MatchCase=defaultNamedOptArg, FieldID=defaultNamedOptArg, TestID=defaultNamedOptArg, SearchAllFields=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(97, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Field
			, Test, Value, Next, MatchCase, FieldID
			, TestID, SearchAllFields)

	def FindFile(self):
		return self._oleobj_.InvokeTypes(2338, LCID, 1, (11, 0), (),)

	def FindNext(self):
		return self._oleobj_.InvokeTypes(2032, LCID, 1, (11, 0), (),)

	def FindPrevious(self):
		return self._oleobj_.InvokeTypes(2033, LCID, 1, (11, 0), (),)

	def FixMe(self):
		return self._oleobj_.InvokeTypes(1323, LCID, 1, (11, 0), (),)

	def FollowHyperlink(self, Address=defaultNamedOptArg, SubAddress=defaultNamedOptArg, AddHistory=defaultNamedOptArg, NewWindow=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1307, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Address
			, SubAddress, AddHistory, NewWindow)

	def Font(self, Name=defaultNamedOptArg, Size=defaultNamedOptArg, Bold=defaultNamedOptArg, Italic=defaultNamedOptArg
			, Underline=defaultNamedOptArg, Color=defaultNamedOptArg, Reset=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(937, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Size, Bold, Italic, Underline, Color
			, Reset)

	def Font32Ex(self, Name=defaultNamedOptArg, Size=defaultNamedOptArg, Bold=defaultNamedOptArg, Italic=defaultNamedOptArg
			, Underline=defaultNamedOptArg, Color=defaultNamedOptArg, Reset=defaultNamedOptArg, CellColor=defaultNamedOptArg, Pattern=defaultNamedOptArg
			, Strikethrough=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2149, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Size, Bold, Italic, Underline, Color
			, Reset, CellColor, Pattern, Strikethrough)

	def FontBold(self, Set=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2320, LCID, 1, (11, 0), ((12, 16),),Set
			)

	def FontEx(self, Name=defaultNamedOptArg, Size=defaultNamedOptArg, Bold=defaultNamedOptArg, Italic=defaultNamedOptArg
			, Underline=defaultNamedOptArg, Color=defaultNamedOptArg, Reset=defaultNamedOptArg, CellColor=defaultNamedOptArg, Pattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2264, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Size, Bold, Italic, Underline, Color
			, Reset, CellColor, Pattern)

	def FontItalic(self, Set=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2321, LCID, 1, (11, 0), ((12, 16),),Set
			)

	def FontStrikethrough(self, Set=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2294, LCID, 1, (11, 0), ((12, 16),),Set
			)

	def FontUnderLine(self, Set=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2324, LCID, 1, (11, 0), ((12, 16),),Set
			)

	def Form(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1004, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def FormViewShow(self):
		return self._oleobj_.InvokeTypes(2074, LCID, 1, (11, 0), (),)

	def FormatCopy(self):
		return self._oleobj_.InvokeTypes(2312, LCID, 1, (11, 0), (),)

	def FormatPainter(self):
		return self._oleobj_.InvokeTypes(2333, LCID, 1, (11, 0), (),)

	def FormatPaste(self):
		return self._oleobj_.InvokeTypes(2313, LCID, 1, (11, 0), (),)

	def GanttBarEditEx(self, Item=defaultNamedNotOptArg, Create=defaultNamedOptArg, Name=defaultNamedOptArg, StartShape=defaultNamedOptArg
			, StartType=defaultNamedOptArg, StartColor=defaultNamedOptArg, MiddleShape=defaultNamedOptArg, MiddleColor=defaultNamedOptArg, MiddlePattern=defaultNamedOptArg
			, EndShape=defaultNamedOptArg, EndType=defaultNamedOptArg, EndColor=defaultNamedOptArg, ShowFor=defaultNamedOptArg, Row=defaultNamedOptArg
			, From=defaultNamedOptArg, To=defaultNamedOptArg, BottomText=defaultNamedOptArg, TopText=defaultNamedOptArg, LeftText=defaultNamedOptArg
			, RightText=defaultNamedOptArg, InsideText=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2145, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, Create, Name, StartShape, StartType, StartColor
			, MiddleShape, MiddleColor, MiddlePattern, EndShape, EndType
			, EndColor, ShowFor, Row, From, To
			, BottomText, TopText, LeftText, RightText, InsideText
			)

	def GanttBarFormat(self, TaskID=defaultNamedOptArg, GanttStyle=defaultNamedOptArg, StartShape=defaultNamedOptArg, StartType=defaultNamedOptArg
			, StartColor=defaultNamedOptArg, MiddleShape=defaultNamedOptArg, MiddlePattern=defaultNamedOptArg, MiddleColor=defaultNamedOptArg, EndShape=defaultNamedOptArg
			, EndType=defaultNamedOptArg, EndColor=defaultNamedOptArg, LeftText=defaultNamedOptArg, RightText=defaultNamedOptArg, TopText=defaultNamedOptArg
			, BottomText=defaultNamedOptArg, InsideText=defaultNamedOptArg, Reset=defaultNamedOptArg, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(938, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),TaskID
			, GanttStyle, StartShape, StartType, StartColor, MiddleShape
			, MiddlePattern, MiddleColor, EndShape, EndType, EndColor
			, LeftText, RightText, TopText, BottomText, InsideText
			, Reset, ProjectName)

	def GanttBarFormatEx(self, TaskID=defaultNamedOptArg, GanttStyle=defaultNamedOptArg, StartShape=defaultNamedOptArg, StartType=defaultNamedOptArg
			, StartColor=defaultNamedOptArg, MiddleShape=defaultNamedOptArg, MiddlePattern=defaultNamedOptArg, MiddleColor=defaultNamedOptArg, EndShape=defaultNamedOptArg
			, EndType=defaultNamedOptArg, EndColor=defaultNamedOptArg, LeftText=defaultNamedOptArg, RightText=defaultNamedOptArg, TopText=defaultNamedOptArg
			, BottomText=defaultNamedOptArg, InsideText=defaultNamedOptArg, Reset=defaultNamedOptArg, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2165, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),TaskID
			, GanttStyle, StartShape, StartType, StartColor, MiddleShape
			, MiddlePattern, MiddleColor, EndShape, EndType, EndColor
			, LeftText, RightText, TopText, BottomText, InsideText
			, Reset, ProjectName)

	def GanttBarLinks(self, Display=0):
		return self._oleobj_.InvokeTypes(2071, LCID, 1, (11, 0), ((3, 48),),Display
			)

	def GanttBarSize(self, Size=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2058, LCID, 1, (11, 0), ((3, 0),),Size
			)

	def GanttBarStyleBaseline(self, Baseline=defaultNamedNotOptArg, Show=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(83, LCID, 1, (11, 0), ((2, 0), (11, 0)),Baseline
			, Show)

	def GanttBarStyleCritical(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(80, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def GanttBarStyleDelete(self, Item=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2059, LCID, 1, (11, 0), ((8, 0),),Item
			)

	def GanttBarStyleEdit(self, Item=defaultNamedNotOptArg, Create=defaultNamedOptArg, Name=defaultNamedOptArg, StartShape=defaultNamedOptArg
			, StartType=defaultNamedOptArg, StartColor=defaultNamedOptArg, MiddleShape=defaultNamedOptArg, MiddleColor=defaultNamedOptArg, MiddlePattern=defaultNamedOptArg
			, EndShape=defaultNamedOptArg, EndType=defaultNamedOptArg, EndColor=defaultNamedOptArg, ShowFor=defaultNamedOptArg, Row=defaultNamedOptArg
			, From=defaultNamedOptArg, To=defaultNamedOptArg, BottomText=defaultNamedOptArg, TopText=defaultNamedOptArg, LeftText=defaultNamedOptArg
			, RightText=defaultNamedOptArg, InsideText=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2060, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, Create, Name, StartShape, StartType, StartColor
			, MiddleShape, MiddleColor, MiddlePattern, EndShape, EndType
			, EndColor, ShowFor, Row, From, To
			, BottomText, TopText, LeftText, RightText, InsideText
			)

	def GanttBarStyleLate(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(82, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def GanttBarStyleSlack(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(81, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def GanttBarStyleSlippage(self, Baseline=defaultNamedNotOptArg, Show=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(84, LCID, 1, (11, 0), ((2, 0), (11, 0)),Baseline
			, Show)

	def GanttBarTextDateFormat(self, DateFormat=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2081, LCID, 1, (11, 0), ((3, 0),),DateFormat
			)

	def GanttChartWizard(self):
		return self._oleobj_.InvokeTypes(2500, LCID, 1, (11, 0), (),)

	def GanttRollup(self, AlwaysRollup=defaultNamedOptArg, HideWhenSummaryExpanded=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2119, LCID, 1, (11, 0), ((12, 16), (12, 16)),AlwaysRollup
			, HideWhenSummaryExpanded)

	def GanttShowBarSplits(self, Display=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2381, LCID, 1, (11, 0), ((12, 16),),Display
			)

	def GanttShowDrawings(self, Display=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2079, LCID, 1, (11, 0), ((12, 16),),Display
			)

	# Result is of type Cell
	def GetCellInfo(self, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(5230, LCID, 1, (9, 0), ((3, 0), (3, 0)),x
			, y)
		if ret is not None:
			ret = Dispatch(ret, u'GetCellInfo', '{00020B19-0000-0000-C000-000000000046}')
		return ret

	def GetCurrentTheme(self):
		return self._oleobj_.InvokeTypes(5211, LCID, 1, (3, 0), (),)

	def GetProjectServerSettings(self, RequestXML=defaultNamedNotOptArg, Project=defaultNamedOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5075, LCID, 1, (8, 0), ((8, 0), (12, 16)),RequestXML
			, Project)

	def GetProjectServerSettingsEx(self):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5078, LCID, 1, (8, 0), (),)

	def GetProjectServerVersion(self, ServerURL=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5080, LCID, 1, (3, 0), ((8, 0),),ServerURL
			)

	def GetRedoListCount(self):
		return self._oleobj_.InvokeTypes(5227, LCID, 1, (3, 0), (),)

	def GetRedoListItem(self, ItemIndex=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5229, LCID, 1, (8, 0), ((3, 0),),ItemIndex
			)

	def GetThemedColor(self, elementType=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5236, LCID, 1, (8, 0), ((3, 0),),elementType
			)

	def GetUndoListCount(self):
		return self._oleobj_.InvokeTypes(5226, LCID, 1, (3, 0), (),)

	def GetUndoListItem(self, ItemIndex=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5228, LCID, 1, (8, 0), ((3, 0),),ItemIndex
			)

	def GoToItemInVersions(self):
		return self._oleobj_.InvokeTypes(2186, LCID, 1, (11, 0), (),)

	def GoalAreaChange(self, goalArea=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(51, LCID, 1, (11, 0), ((2, 0),),goalArea
			)

	def GoalAreaHighlight(self, goalArea=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5063, LCID, 1, (24, 0), ((3, 0),),goalArea
			)

	def GoalAreaTaskHighlight(self, TaskID=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5065, LCID, 1, (24, 0), ((3, 0),),TaskID
			)

	def GotoNextOverAllocation(self):
		return self._oleobj_.InvokeTypes(214, LCID, 1, (11, 0), (),)

	def GotoTaskDates(self):
		return self._oleobj_.InvokeTypes(2054, LCID, 1, (11, 0), (),)

	def Gridlines(self):
		return self._oleobj_.InvokeTypes(912, LCID, 1, (11, 0), (),)

	def GridlinesEdit(self, Item=defaultNamedNotOptArg, NormalType=defaultNamedOptArg, NormalColor=defaultNamedOptArg, Interval=defaultNamedOptArg
			, IntervalType=defaultNamedOptArg, IntervalColor=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2061, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, NormalType, NormalColor, Interval, IntervalType, IntervalColor
			)

	def GridlinesEditEx(self, Item=defaultNamedNotOptArg, NormalType=defaultNamedOptArg, NormalColor=defaultNamedOptArg, Interval=defaultNamedOptArg
			, IntervalType=defaultNamedOptArg, IntervalColor=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2152, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, NormalType, NormalColor, Interval, IntervalType, IntervalColor
			)

	def GroupApply(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(512, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def GroupBy(self):
		return self._oleobj_.InvokeTypes(513, LCID, 1, (11, 0), (),)

	def GroupClear(self):
		return self._oleobj_.InvokeTypes(515, LCID, 1, (11, 0), (),)

	def GroupMaintainHierarchy(self, On=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2296, LCID, 1, (11, 0), ((11, 0),),On
			)

	def GroupNew(self):
		return self._oleobj_.InvokeTypes(514, LCID, 1, (11, 0), (),)

	def Groups(self):
		return self._oleobj_.InvokeTypes(511, LCID, 1, (11, 0), (),)

	def HelpAbout(self):
		return self._oleobj_.InvokeTypes(806, LCID, 1, (11, 0), (),)

	def HelpAnswerWizard(self):
		return self._oleobj_.InvokeTypes(816, LCID, 1, (11, 0), (),)

	def HelpContents(self):
		return self._oleobj_.InvokeTypes(804, LCID, 1, (11, 0), (),)

	def HelpContextHelp(self):
		return self._oleobj_.InvokeTypes(814, LCID, 1, (11, 0), (),)

	def HelpCreateYourProject(self):
		return self._oleobj_.InvokeTypes(2374, LCID, 1, (11, 0), (),)

	def HelpCueCards(self, Filename=defaultNamedOptArg, ContextNumber=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(813, LCID, 1, (11, 0), ((12, 16), (12, 16)),Filename
			, ContextNumber)

	def HelpLaunch(self, Filename=defaultNamedOptArg, ContextNumber=defaultNamedOptArg, Search=defaultNamedOptArg, SearchKey=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(810, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Filename
			, ContextNumber, Search, SearchKey)

	def HelpMSProjectFundamentals(self):
		return self._oleobj_.InvokeTypes(817, LCID, 1, (11, 0), (),)

	def HelpOnlineIndex(self):
		return self._oleobj_.InvokeTypes(809, LCID, 1, (11, 0), (),)

	def HelpQuickPreview(self):
		return self._oleobj_.InvokeTypes(811, LCID, 1, (11, 0), (),)

	def HelpReference(self):
		return self._oleobj_.InvokeTypes(2127, LCID, 1, (11, 0), (),)

	def HelpSearch(self):
		return self._oleobj_.InvokeTypes(808, LCID, 1, (11, 0), (),)

	def HelpTechnicalSupport(self):
		return self._oleobj_.InvokeTypes(812, LCID, 1, (11, 0), (),)

	def HelpTopics(self):
		return self._oleobj_.InvokeTypes(815, LCID, 1, (11, 0), (),)

	def HelpWhatsNew(self):
		return self._oleobj_.InvokeTypes(2129, LCID, 1, (11, 0), (),)

	def ImportCommitment(self, CommitmentDate=defaultNamedOptArg, CommitmentGuid=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2098, LCID, 1, (11, 0), ((12, 16), (12, 16)),CommitmentDate
			, CommitmentGuid)

	def ImportOutlookTasks(self):
		return self._oleobj_.InvokeTypes(2121, LCID, 1, (11, 0), (),)

	def ImportResourceList(self, ServerURL=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2122, LCID, 1, (11, 0), ((12, 16),),ServerURL
			)

	def InactivateTaskToggle(self, MakeActive=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(91, LCID, 1, (11, 0), ((12, 16),),MakeActive
			)

	def InformationDialog(self, Tab=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(217, LCID, 1, (11, 0), ((12, 16),),Tab
			)

	def InsertBlankRow(self):
		return self._oleobj_.InvokeTypes(2171, LCID, 1, (11, 0), (),)

	def InsertHyperlink(self, Name=defaultNamedOptArg, Address=defaultNamedOptArg, SubAddress=defaultNamedOptArg, ScreenTip=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1309, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Address, SubAddress, ScreenTip)

	def InsertManualTask(self):
		return self._oleobj_.InvokeTypes(2169, LCID, 1, (11, 0), (),)

	def InsertMilestoneTask(self):
		return self._oleobj_.InvokeTypes(2170, LCID, 1, (11, 0), (),)

	def InsertNotes(self):
		return self._oleobj_.InvokeTypes(2078, LCID, 1, (11, 0), (),)

	def InsertResource(self, Type=0):
		return self._oleobj_.InvokeTypes(2179, LCID, 1, (11, 0), ((3, 48),),Type
			)

	def InsertScheduledTask(self):
		return self._oleobj_.InvokeTypes(2168, LCID, 1, (11, 0), (),)

	def InsertSummaryTask(self):
		return self._oleobj_.InvokeTypes(2180, LCID, 1, (11, 0), (),)

	def InsertTask(self):
		return self._oleobj_.InvokeTypes(2167, LCID, 1, (11, 0), (),)

	def IsCommandEnabled(self, CommandName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5237, LCID, 1, (3, 0), ((8, 0),),CommandName
			)

	def IsOfficeTaskPaneVisible(self):
		return self._oleobj_.InvokeTypes(5209, LCID, 1, (11, 0), (),)

	def IsOffline(self):
		return self._oleobj_.InvokeTypes(5233, LCID, 1, (11, 0), (),)

	def IsReducedFunctionalityMode(self):
		return self._oleobj_.InvokeTypes(5234, LCID, 1, (11, 0), (),)

	def IsURLTrusted(self, URL=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5086, LCID, 1, (11, 0), ((8, 0),),URL
			)

	def IsUndoingOrRedoing(self):
		return self._oleobj_.InvokeTypes(5216, LCID, 1, (11, 0), (),)

	def Layout(self):
		return self._oleobj_.InvokeTypes(941, LCID, 1, (11, 0), (),)

	def LayoutNow(self):
		return self._oleobj_.InvokeTypes(907, LCID, 1, (11, 0), (),)

	def LayoutRelatedNow(self):
		return self._oleobj_.InvokeTypes(30, LCID, 1, (11, 0), (),)

	def LayoutSelectionNow(self):
		return self._oleobj_.InvokeTypes(2399, LCID, 1, (11, 0), (),)

	def LevelNow(self, All=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(609, LCID, 1, (11, 0), ((12, 16),),All
			)

	def LevelSelected(self, ResolveMethod=0):
		return self._oleobj_.InvokeTypes(2292, LCID, 1, (11, 0), ((3, 48),),ResolveMethod
			)

	def LevelingClear(self, All=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(612, LCID, 1, (11, 0), ((12, 16),),All
			)

	def LevelingOptions(self, Automatic=defaultNamedOptArg, DelayInSlack=defaultNamedOptArg, AutoClearLeveling=defaultNamedOptArg, Order=defaultNamedOptArg
			, LevelEntireProject=defaultNamedOptArg, FromDate=defaultNamedOptArg, ToDate=defaultNamedOptArg, PeriodBasis=defaultNamedOptArg, LevelIndividualAssignments=defaultNamedOptArg
			, LevelingCanSplit=defaultNamedOptArg, LevelProposedBookings=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(608, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Automatic
			, DelayInSlack, AutoClearLeveling, Order, LevelEntireProject, FromDate
			, ToDate, PeriodBasis, LevelIndividualAssignments, LevelingCanSplit, LevelProposedBookings
			)

	def LevelingOptionsEx(self, Automatic=defaultNamedOptArg, DelayInSlack=defaultNamedOptArg, AutoClearLeveling=defaultNamedOptArg, Order=defaultNamedOptArg
			, LevelEntireProject=defaultNamedOptArg, FromDate=defaultNamedOptArg, ToDate=defaultNamedOptArg, PeriodBasis=defaultNamedOptArg, LevelIndividualAssignments=defaultNamedOptArg
			, LevelingCanSplit=defaultNamedOptArg, LevelProposedBookings=defaultNamedOptArg, LevelPinnedTasks=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2249, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Automatic
			, DelayInSlack, AutoClearLeveling, Order, LevelEntireProject, FromDate
			, ToDate, PeriodBasis, LevelIndividualAssignments, LevelingCanSplit, LevelProposedBookings
			, LevelPinnedTasks)

	def LinkTasks(self):
		return self._oleobj_.InvokeTypes(210, LCID, 1, (11, 0), (),)

	def LinkTasksEdit(self, From=defaultNamedNotOptArg, To=defaultNamedNotOptArg, Delete=defaultNamedOptArg, Type=defaultNamedOptArg
			, Lag=defaultNamedOptArg, PredecessorProjectName=defaultNamedOptArg, SuccessorProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2052, LCID, 1, (11, 0), ((3, 0), (3, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),From
			, To, Delete, Type, Lag, PredecessorProjectName
			, SuccessorProjectName)

	def LinksBetweenProjects(self, AcceptAll=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(245, LCID, 1, (11, 0), ((12, 16),),AcceptAll
			)

	def LoadTasksFromServer(self, strURL=defaultNamedNotOptArg, strListGuid=defaultNamedNotOptArg, strListName=defaultNamedNotOptArg, bstrViewGUID=defaultNamedNotOptArg
			, bstrViewName=defaultNamedNotOptArg, iListID=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5238, LCID, 1, (8, 0), ((8, 0), (8, 0), (8, 0), (8, 0), (8, 0), (3, 0)),strURL
			, strListGuid, strListName, bstrViewGUID, bstrViewName, iListID
			)

	def LoadWebBrowserControl(self, TargetPage=defaultNamedNotOptArg, WrapperPage=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(54, LCID, 1, (11, 0), ((8, 0), (12, 16)),TargetPage
			, WrapperPage)

	def LoadWebBrowserControlEx(self, TargetPage=defaultNamedNotOptArg, WrapperPage=defaultNamedOptArg, FunctionalityName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2271, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16)),TargetPage
			, WrapperPage, FunctionalityName)

	def LoadWebPaneControl(self, TargetPage=defaultNamedNotOptArg, WrapperPage=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(56, LCID, 1, (11, 0), ((8, 0), (12, 16)),TargetPage
			, WrapperPage)

	def LocaleID(self):
		return self._oleobj_.InvokeTypes(5084, LCID, 1, (3, 0), (),)

	def LookUpTableAdd(self, FieldID=defaultNamedNotOptArg, Level=defaultNamedOptArg, Code=defaultNamedOptArg, Description=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(635, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16)),FieldID
			, Level, Code, Description)

	def LookUpTableAddEx(self, FieldID=defaultNamedNotOptArg, Level=defaultNamedOptArg, Code=defaultNamedOptArg, Description=defaultNamedOptArg
			, Phonetic=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2266, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16), (12, 16), (12, 16)),FieldID
			, Level, Code, Description, Phonetic)

	def Macro(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1001, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def MacroSecurity(self):
		return self._oleobj_.InvokeTypes(2396, LCID, 1, (11, 0), (),)

	def MacroShowCode(self):
		return self._oleobj_.InvokeTypes(2246, LCID, 1, (11, 0), (),)

	def MacroShowVba(self):
		return self._oleobj_.InvokeTypes(2245, LCID, 1, (11, 0), (),)

	def MailLogoff(self):
		return self._oleobj_.InvokeTypes(5035, LCID, 1, (24, 0), (),)

	def MailLogon(self, Name=defaultNamedOptArg, Password=defaultNamedOptArg, DownloadNewMail=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(5034, LCID, 1, (24, 0), ((12, 16), (12, 16), (12, 16)),Name
			, Password, DownloadNewMail)

	def MailOpen(self, Any=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(119, LCID, 1, (11, 0), ((12, 16),),Any
			)

	def MailPostDocument(self):
		return self._oleobj_.InvokeTypes(131, LCID, 1, (11, 0), (),)

	def MailProjectMailCustomize(self, action=defaultNamedOptArg, Position=defaultNamedOptArg, FieldID=defaultNamedOptArg, Title=defaultNamedOptArg
			, IncludeInTeamStatus=defaultNamedOptArg, Editable=defaultNamedOptArg, UseAssignmentField=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2373, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),action
			, Position, FieldID, Title, IncludeInTeamStatus, Editable
			, UseAssignmentField)

	def MailRoutingSlip(self, To=defaultNamedOptArg, Subject=defaultNamedOptArg, Body=defaultNamedOptArg, AllAtOnce=defaultNamedOptArg
			, ReturnWhenDone=defaultNamedOptArg, TrackStatus=defaultNamedOptArg, Clear=defaultNamedOptArg, SendNow=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(125, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),To
			, Subject, Body, AllAtOnce, ReturnWhenDone, TrackStatus
			, Clear, SendNow)

	def MailSend(self, To=defaultNamedOptArg, Cc=defaultNamedOptArg, Subject=defaultNamedOptArg, Body=defaultNamedOptArg
			, Enclosures=defaultNamedOptArg, IncludeDocument=defaultNamedOptArg, ReturnReceipt=defaultNamedOptArg, Bcc=defaultNamedOptArg, Urgent=defaultNamedOptArg
			, SaveCopy=defaultNamedOptArg, AddRecipient=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(120, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),To
			, Cc, Subject, Body, Enclosures, IncludeDocument
			, ReturnReceipt, Bcc, Urgent, SaveCopy, AddRecipient
			)

	def MailSendProjectMail(self, MessageType=defaultNamedOptArg, Subject=defaultNamedOptArg, Body=defaultNamedOptArg, Fields=defaultNamedOptArg
			, UpdateAsOf=defaultNamedOptArg, ShowDialog=defaultNamedOptArg, InstallationMessage=defaultNamedOptArg, UpdateFrom=defaultNamedOptArg, PublishScope=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2351, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),MessageType
			, Subject, Body, Fields, UpdateAsOf, ShowDialog
			, InstallationMessage, UpdateFrom, PublishScope)

	def MailSendScheduleNote(self, Manager=defaultNamedOptArg, Resources=defaultNamedOptArg, TaskContacts=defaultNamedOptArg, Selection=defaultNamedOptArg
			, IncludeDocument=defaultNamedOptArg, IncludePicture=defaultNamedOptArg, Body=defaultNamedOptArg, Subject=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(129, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Manager
			, Resources, TaskContacts, Selection, IncludeDocument, IncludePicture
			, Body, Subject)

	def MailSession(self):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5036, LCID, 1, (8, 0), (),)

	def MailSystem(self):
		return self._oleobj_.InvokeTypes(5037, LCID, 1, (3, 0), (),)

	def MailUpdateProject(self, DataFile=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2372, LCID, 1, (11, 0), ((8, 0),),DataFile
			)

	def MakeFieldEnterprise(self, FieldID=defaultNamedOptArg, FieldName=defaultNamedOptArg, LookupTableName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2275, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),FieldID
			, FieldName, LookupTableName)

	def MakeLocalCalendarEnterprise(self, OldName=defaultNamedOptArg, NewName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2369, LCID, 1, (11, 0), ((12, 16), (12, 16)),OldName
			, NewName)

	def ManageSiteColumns(self):
		return self._oleobj_.InvokeTypes(2288, LCID, 1, (11, 0), (),)

	def MapEdit(self, Name=defaultNamedOptArg, Create=defaultNamedOptArg, OverwriteExisting=defaultNamedOptArg, NewName=defaultNamedOptArg
			, DataCategory=defaultNamedOptArg, CategoryEnabled=defaultNamedOptArg, TableName=defaultNamedOptArg, FieldName=defaultNamedOptArg, ExternalFieldName=defaultNamedOptArg
			, ExportFilter=defaultNamedOptArg, ImportMethod=defaultNamedOptArg, MergeKey=defaultNamedOptArg, HeaderRow=defaultNamedOptArg, AssignmentData=defaultNamedOptArg
			, TextDelimiter=defaultNamedOptArg, TextFileOrigin=defaultNamedOptArg, UseHtmlTemplate=defaultNamedOptArg, TemplateFile=defaultNamedOptArg, IncludeImage=defaultNamedOptArg
			, ImageFile=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(243, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Create, OverwriteExisting, NewName, DataCategory, CategoryEnabled
			, TableName, FieldName, ExternalFieldName, ExportFilter, ImportMethod
			, MergeKey, HeaderRow, AssignmentData, TextDelimiter, TextFileOrigin
			, UseHtmlTemplate, TemplateFile, IncludeImage, ImageFile)

	def MenuBarApply(self, Name=defaultNamedNotOptArg, Default=defaultNamedOptArg, NoFiles=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2076, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16)),Name
			, Default, NoFiles)

	def MenuBarEdit(self, Copy=defaultNamedOptArg, Create=defaultNamedOptArg, Name=defaultNamedOptArg, NewName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2075, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Copy
			, Create, Name, NewName)

	def MenuBars(self):
		return self._oleobj_.InvokeTypes(2072, LCID, 1, (11, 0), (),)

	def Message(self, Message=defaultNamedNotOptArg, Type=0, YesText=defaultNamedOptArg, NoText=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2, LCID, 1, (11, 0), ((8, 0), (3, 48), (12, 16), (12, 16)),Message
			, Type, YesText, NoText)

	def NewTasksStartOn(self, StartOnDate=0):
		return self._oleobj_.InvokeTypes(2295, LCID, 1, (11, 0), ((3, 48),),StartOnDate
			)

	def ODBCCreateDataSource(self):
		return self._oleobj_.InvokeTypes(133, LCID, 1, (11, 0), (),)

	def ODBCManageDataSources(self):
		return self._oleobj_.InvokeTypes(132, LCID, 1, (11, 0), (),)

	def ObjectChangeIcon(self):
		return self._oleobj_.InvokeTypes(235, LCID, 1, (11, 0), (),)

	def ObjectConvert(self):
		return self._oleobj_.InvokeTypes(236, LCID, 1, (11, 0), (),)

	def ObjectInsert(self):
		return self._oleobj_.InvokeTypes(221, LCID, 1, (11, 0), (),)

	def ObjectLinks(self):
		return self._oleobj_.InvokeTypes(238, LCID, 1, (11, 0), (),)

	def ObjectVerb(self, Verb=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(237, LCID, 1, (11, 0), ((12, 16),),Verb
			)

	def OfficeOnTheWeb(self):
		return self._oleobj_.InvokeTypes(1322, LCID, 1, (11, 0), (),)

	def OfficeTaskPaneHide(self):
		return self._oleobj_.InvokeTypes(5210, LCID, 1, (24, 0), (),)

	def OpenFromSharePoint(self, SiteURL=defaultNamedOptArg, ListName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2293, LCID, 1, (11, 0), ((12, 16), (12, 16)),SiteURL
			, ListName)

	def OpenServerPage(self, Page=0):
		return self._oleobj_.InvokeTypes(636, LCID, 1, (11, 0), ((3, 48),),Page
			)

	def OpenUndoTransaction(self, Label=defaultNamedNotOptArg, Guid=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(5219, LCID, 1, (24, 0), ((8, 0), (12, 16)),Label
			, Guid)

	def OpenXML(self, XML=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5079, LCID, 1, (3, 0), ((8, 0),),XML
			)

	def OptionsCalculation(self, Automatic=defaultNamedOptArg, AutoTrack=defaultNamedOptArg, SpreadPercentToStatusDate=defaultNamedOptArg, SpreadCostsToStatusDate=defaultNamedOptArg
			, AutoCalcCosts=defaultNamedOptArg, FixedCostAccrual=defaultNamedOptArg, CalcMultipleCriticalPaths=defaultNamedOptArg, CriticalSlack=defaultNamedOptArg, SetDefaults=defaultNamedOptArg
			, CalcInsProjLikeSummTask=defaultNamedOptArg, MoveCompleted=defaultNamedOptArg, AndMoveRemaining=defaultNamedOptArg, MoveRemaining=defaultNamedOptArg, AndMoveCompleted=defaultNamedOptArg
			, EVMethod=defaultNamedOptArg, EVBaseline=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(606, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Automatic
			, AutoTrack, SpreadPercentToStatusDate, SpreadCostsToStatusDate, AutoCalcCosts, FixedCostAccrual
			, CalcMultipleCriticalPaths, CriticalSlack, SetDefaults, CalcInsProjLikeSummTask, MoveCompleted
			, AndMoveRemaining, MoveRemaining, AndMoveCompleted, EVMethod, EVBaseline
			)

	def OptionsCalendar(self, StartWeekOnMonday=defaultNamedOptArg, StartYearIn=defaultNamedOptArg, StartTime=defaultNamedOptArg, FinishTime=defaultNamedOptArg
			, HoursPerDay=defaultNamedOptArg, HoursPerWeek=defaultNamedOptArg, SetDefaults=defaultNamedOptArg, StartWeekOn=defaultNamedOptArg, UseFYStartYear=defaultNamedOptArg
			, DaysPerMonth=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(649, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),StartWeekOnMonday
			, StartYearIn, StartTime, FinishTime, HoursPerDay, HoursPerWeek
			, SetDefaults, StartWeekOn, UseFYStartYear, DaysPerMonth)

	def OptionsEdit(self, MoveAfterReturn=defaultNamedOptArg, DragAndDrop=defaultNamedOptArg, UpdateLinks=defaultNamedOptArg, CopyResourceUsageHeader=defaultNamedOptArg
			, PhoneticInfo=defaultNamedOptArg, PhoneticType=defaultNamedOptArg, MinuteLabelDisplay=defaultNamedOptArg, HourLabelDisplay=defaultNamedOptArg, DayLabelDisplay=defaultNamedOptArg
			, WeekLabelDisplay=defaultNamedOptArg, YearLabelDisplay=defaultNamedOptArg, SpaceBeforeTimeLabel=defaultNamedOptArg, SetDefaults=defaultNamedOptArg, MonthLabelDisplay=defaultNamedOptArg
			, SetDefaultsTimeUnits=defaultNamedOptArg, HyperlinkColor=defaultNamedOptArg, FollowedHyperlinkColor=defaultNamedOptArg, UnderlineHyperlinks=defaultNamedOptArg, SetDefaultsHyperlink=defaultNamedOptArg
			, InCellEditing=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(641, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),MoveAfterReturn
			, DragAndDrop, UpdateLinks, CopyResourceUsageHeader, PhoneticInfo, PhoneticType
			, MinuteLabelDisplay, HourLabelDisplay, DayLabelDisplay, WeekLabelDisplay, YearLabelDisplay
			, SpaceBeforeTimeLabel, SetDefaults, MonthLabelDisplay, SetDefaultsTimeUnits, HyperlinkColor
			, FollowedHyperlinkColor, UnderlineHyperlinks, SetDefaultsHyperlink, InCellEditing)

	def OptionsEditEx(self, MoveAfterReturn=defaultNamedOptArg, DragAndDrop=defaultNamedOptArg, UpdateLinks=defaultNamedOptArg, CopyResourceUsageHeader=defaultNamedOptArg
			, PhoneticInfo=defaultNamedOptArg, PhoneticType=defaultNamedOptArg, MinuteLabelDisplay=defaultNamedOptArg, HourLabelDisplay=defaultNamedOptArg, DayLabelDisplay=defaultNamedOptArg
			, WeekLabelDisplay=defaultNamedOptArg, YearLabelDisplay=defaultNamedOptArg, SpaceBeforeTimeLabel=defaultNamedOptArg, SetDefaults=defaultNamedOptArg, MonthLabelDisplay=defaultNamedOptArg
			, SetDefaultsTimeUnits=defaultNamedOptArg, HyperlinkColor=defaultNamedOptArg, FollowedHyperlinkColor=defaultNamedOptArg, UnderlineHyperlinks=defaultNamedOptArg, SetDefaultsHyperlink=defaultNamedOptArg
			, InCellEditing=defaultNamedOptArg, AllowTaskDelegation=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2159, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),MoveAfterReturn
			, DragAndDrop, UpdateLinks, CopyResourceUsageHeader, PhoneticInfo, PhoneticType
			, MinuteLabelDisplay, HourLabelDisplay, DayLabelDisplay, WeekLabelDisplay, YearLabelDisplay
			, SpaceBeforeTimeLabel, SetDefaults, MonthLabelDisplay, SetDefaultsTimeUnits, HyperlinkColor
			, FollowedHyperlinkColor, UnderlineHyperlinks, SetDefaultsHyperlink, InCellEditing, AllowTaskDelegation
			)

	def OptionsGeneral(self, PlanningWizard=defaultNamedOptArg, WizardUsage=defaultNamedOptArg, WizardErrors=defaultNamedOptArg, WizardScheduling=defaultNamedOptArg
			, ShowTipOfDay=defaultNamedOptArg, AutoAddResources=defaultNamedOptArg, StandardRate=defaultNamedOptArg, OvertimeRate=defaultNamedOptArg, LastFile=defaultNamedOptArg
			, SummaryInfo=defaultNamedOptArg, UserName=defaultNamedOptArg, SetDefaults=defaultNamedOptArg, ShowWelcome=defaultNamedOptArg, AutoFilter=defaultNamedOptArg
			, MacroVirusProtection=defaultNamedOptArg, DisplayRecentFiles=defaultNamedOptArg, RecentFilesMaximum=defaultNamedOptArg, FontConversion=defaultNamedOptArg, ShowStartupWorkpane=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(642, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),PlanningWizard
			, WizardUsage, WizardErrors, WizardScheduling, ShowTipOfDay, AutoAddResources
			, StandardRate, OvertimeRate, LastFile, SummaryInfo, UserName
			, SetDefaults, ShowWelcome, AutoFilter, MacroVirusProtection, DisplayRecentFiles
			, RecentFilesMaximum, FontConversion, ShowStartupWorkpane)

	def OptionsGeneralEx(self, PlanningWizard=defaultNamedOptArg, WizardUsage=defaultNamedOptArg, WizardErrors=defaultNamedOptArg, WizardScheduling=defaultNamedOptArg
			, ShowTipOfDay=defaultNamedOptArg, AutoAddResources=defaultNamedOptArg, StandardRate=defaultNamedOptArg, OvertimeRate=defaultNamedOptArg, LastFile=defaultNamedOptArg
			, SummaryInfo=defaultNamedOptArg, UserName=defaultNamedOptArg, SetDefaults=defaultNamedOptArg, ShowWelcome=defaultNamedOptArg, AutoFilter=defaultNamedOptArg
			, MacroVirusProtection=defaultNamedOptArg, DisplayRecentFiles=defaultNamedOptArg, RecentFilesMaximum=defaultNamedOptArg, FontConversion=defaultNamedOptArg, ShowStartupWorkpane=defaultNamedOptArg
			, MaxUndoRecords=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2261, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),PlanningWizard
			, WizardUsage, WizardErrors, WizardScheduling, ShowTipOfDay, AutoAddResources
			, StandardRate, OvertimeRate, LastFile, SummaryInfo, UserName
			, SetDefaults, ShowWelcome, AutoFilter, MacroVirusProtection, DisplayRecentFiles
			, RecentFilesMaximum, FontConversion, ShowStartupWorkpane, MaxUndoRecords)

	def OptionsInterface(self, ShowResourceAssignmentIndicators=defaultNamedOptArg, ShowEditToStartFinishDates=defaultNamedOptArg, ShowEditsToWorkUnitsDurationIndicators=defaultNamedOptArg, ShowDeletionInNameColumn=defaultNamedOptArg
			, DisplayProjectGuide=defaultNamedOptArg, ProjectGuideUseDefaultFunctionalLayoutPage=defaultNamedOptArg, ProjectGuideFunctionalLayoutPage=defaultNamedOptArg, ProjectGuideUseDefaultContent=defaultNamedOptArg, ProjectGuideContent=defaultNamedOptArg
			, SetAsDefaults=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(651, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),ShowResourceAssignmentIndicators
			, ShowEditToStartFinishDates, ShowEditsToWorkUnitsDurationIndicators, ShowDeletionInNameColumn, DisplayProjectGuide, ProjectGuideUseDefaultFunctionalLayoutPage
			, ProjectGuideFunctionalLayoutPage, ProjectGuideUseDefaultContent, ProjectGuideContent, SetAsDefaults)

	def OptionsInterfaceEx(self, ShowResourceAssignmentIndicators=defaultNamedOptArg, ShowEditToStartFinishDates=defaultNamedOptArg, ShowEditsToWorkUnitsDurationIndicators=defaultNamedOptArg, ShowDeletionInNameColumn=defaultNamedOptArg
			, DisplayProjectGuide=defaultNamedOptArg, ProjectGuideUseDefaultFunctionalLayoutPage=defaultNamedOptArg, ProjectGuideFunctionalLayoutPage=defaultNamedOptArg, ProjectGuideUseDefaultContent=defaultNamedOptArg, ProjectGuideContent=defaultNamedOptArg
			, SetAsDefaults=defaultNamedOptArg, UseOMIDs=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2268, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),ShowResourceAssignmentIndicators
			, ShowEditToStartFinishDates, ShowEditsToWorkUnitsDurationIndicators, ShowDeletionInNameColumn, DisplayProjectGuide, ProjectGuideUseDefaultFunctionalLayoutPage
			, ProjectGuideFunctionalLayoutPage, ProjectGuideUseDefaultContent, ProjectGuideContent, SetAsDefaults, UseOMIDs
			)

	def OptionsPreferences(self):
		return self._oleobj_.InvokeTypes(603, LCID, 1, (11, 0), (),)

	def OptionsSave(self, DefaultSaveFormat=defaultNamedOptArg, DefaultProjectsPath=defaultNamedOptArg, DefaultUserTemplatesPath=defaultNamedOptArg, DefaultWorkgroupTemplatesPath=defaultNamedOptArg
			, ExpandDatabaseTimephasedData=defaultNamedOptArg, AutomaticSave=defaultNamedOptArg, AutomaticSaveInterval=defaultNamedOptArg, AutomaticSaveOptions=defaultNamedOptArg, AutomaticSavePrompt=defaultNamedOptArg
			, SetDefaultsDatabase=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(650, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),DefaultSaveFormat
			, DefaultProjectsPath, DefaultUserTemplatesPath, DefaultWorkgroupTemplatesPath, ExpandDatabaseTimephasedData, AutomaticSave
			, AutomaticSaveInterval, AutomaticSaveOptions, AutomaticSavePrompt, SetDefaultsDatabase)

	def OptionsSchedule(self, ScheduleMessages=defaultNamedOptArg, StartOnCurrentDate=defaultNamedOptArg, AutoLink=defaultNamedOptArg, AutoSplit=defaultNamedOptArg
			, CriticalSlack=defaultNamedOptArg, TaskType=defaultNamedOptArg, DurationUnits=defaultNamedOptArg, WorkUnits=defaultNamedOptArg, AutoTrack=defaultNamedOptArg
			, SetDefaults=defaultNamedOptArg, AssignmentUnits=defaultNamedOptArg, EffortDriven=defaultNamedOptArg, HonorConstraints=defaultNamedOptArg, ShowEstimated=defaultNamedOptArg
			, NewTasksEstimated=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(644, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),ScheduleMessages
			, StartOnCurrentDate, AutoLink, AutoSplit, CriticalSlack, TaskType
			, DurationUnits, WorkUnits, AutoTrack, SetDefaults, AssignmentUnits
			, EffortDriven, HonorConstraints, ShowEstimated, NewTasksEstimated)

	def OptionsSecurity(self, RemoveFileProperties=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(652, LCID, 1, (11, 0), ((12, 16),),RemoveFileProperties
			)

	def OptionsSecurityEx(self, RemoveFileProperties=defaultNamedOptArg, TrustWSS=defaultNamedOptArg, LegacyFileFormats=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2272, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),RemoveFileProperties
			, TrustWSS, LegacyFileFormats)

	def OptionsSecurityTab(self, DefaultTab=0):
		return self._oleobj_.InvokeTypes(2504, LCID, 1, (11, 0), ((3, 48),),DefaultTab
			)

	def OptionsSpelling(self, TaskName=defaultNamedOptArg, TaskNotes=defaultNamedOptArg, TaskText1=defaultNamedOptArg, TaskText2=defaultNamedOptArg
			, TaskText3=defaultNamedOptArg, TaskText4=defaultNamedOptArg, TaskText5=defaultNamedOptArg, TaskText6=defaultNamedOptArg, TaskText7=defaultNamedOptArg
			, TaskText8=defaultNamedOptArg, TaskText9=defaultNamedOptArg, TaskText10=defaultNamedOptArg, ResourceCode=defaultNamedOptArg, ResourceName=defaultNamedOptArg
			, ResourceNotes=defaultNamedOptArg, ResourceGroup=defaultNamedOptArg, ResourceText1=defaultNamedOptArg, ResourceText2=defaultNamedOptArg, ResourceText3=defaultNamedOptArg
			, ResourceText4=defaultNamedOptArg, ResourceText5=defaultNamedOptArg, AssignNotes=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg, IgnoreNumberWords=defaultNamedOptArg
			, AlwaysSuggest=defaultNamedOptArg, UseCustomDictionary=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(614, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),TaskName
			, TaskNotes, TaskText1, TaskText2, TaskText3, TaskText4
			, TaskText5, TaskText6, TaskText7, TaskText8, TaskText9
			, TaskText10, ResourceCode, ResourceName, ResourceNotes, ResourceGroup
			, ResourceText1, ResourceText2, ResourceText3, ResourceText4, ResourceText5
			, AssignNotes, IgnoreUppercase, IgnoreNumberWords, AlwaysSuggest, UseCustomDictionary
			)

	def OptionsView(self, DefaultView=defaultNamedOptArg, DateFormat=defaultNamedOptArg, ProjectSummary=defaultNamedOptArg, DisplayStatusBar=defaultNamedOptArg
			, DisplayEntryBar=defaultNamedOptArg, DisplayScrollBars=defaultNamedOptArg, CurrencySymbol=defaultNamedOptArg, SymbolPlacement=defaultNamedOptArg, CurrencyDigits=defaultNamedOptArg
			, ProjectCurrency=defaultNamedOptArg, DisplayOutlineNumber=defaultNamedOptArg, DisplayOutlineSymbols=defaultNamedOptArg, DisplayNameIndent=defaultNamedOptArg, DisplaySummaryTasks=defaultNamedOptArg
			, DisplayOLEIndicator=defaultNamedOptArg, DisplayExternalSuccessors=defaultNamedOptArg, DisplayExternalPredecessors=defaultNamedOptArg, CrossProjectLinksInfo=defaultNamedOptArg, AcceptNewExternalData=defaultNamedOptArg
			, DisplayWindowsInTaskbar=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(646, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),DefaultView
			, DateFormat, ProjectSummary, DisplayStatusBar, DisplayEntryBar, DisplayScrollBars
			, CurrencySymbol, SymbolPlacement, CurrencyDigits, ProjectCurrency, DisplayOutlineNumber
			, DisplayOutlineSymbols, DisplayNameIndent, DisplaySummaryTasks, DisplayOLEIndicator, DisplayExternalSuccessors
			, DisplayExternalPredecessors, CrossProjectLinksInfo, AcceptNewExternalData, DisplayWindowsInTaskbar)

	def OptionsViewEx(self, DefaultView=defaultNamedOptArg, DateFormat=defaultNamedOptArg, ProjectSummary=defaultNamedOptArg, DisplayStatusBar=defaultNamedOptArg
			, DisplayEntryBar=defaultNamedOptArg, DisplayScrollBars=defaultNamedOptArg, CurrencySymbol=defaultNamedOptArg, SymbolPlacement=defaultNamedOptArg, CurrencyDigits=defaultNamedOptArg
			, ProjectCurrency=defaultNamedOptArg, DisplayOutlineNumber=defaultNamedOptArg, DisplayOutlineSymbols=defaultNamedOptArg, DisplayNameIndent=defaultNamedOptArg, DisplaySummaryTasks=defaultNamedOptArg
			, DisplayOLEIndicator=defaultNamedOptArg, DisplayExternalSuccessors=defaultNamedOptArg, DisplayExternalPredecessors=defaultNamedOptArg, CrossProjectLinksInfo=defaultNamedOptArg, AcceptNewExternalData=defaultNamedOptArg
			, DisplayWindowsInTaskbar=defaultNamedOptArg, DisplayScreentips=defaultNamedOptArg, CalendarType=defaultNamedOptArg, Use3DLook=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2262, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),DefaultView
			, DateFormat, ProjectSummary, DisplayStatusBar, DisplayEntryBar, DisplayScrollBars
			, CurrencySymbol, SymbolPlacement, CurrencyDigits, ProjectCurrency, DisplayOutlineNumber
			, DisplayOutlineSymbols, DisplayNameIndent, DisplaySummaryTasks, DisplayOLEIndicator, DisplayExternalSuccessors
			, DisplayExternalPredecessors, CrossProjectLinksInfo, AcceptNewExternalData, DisplayWindowsInTaskbar, DisplayScreentips
			, CalendarType, Use3DLook)

	def OptionsWorkgroup(self, WorkgroupMessages=defaultNamedOptArg, ServerURL=defaultNamedOptArg, ServerPath=defaultNamedOptArg, ReceiveNotifications=defaultNamedOptArg
			, SendHyperlinkNote=defaultNamedOptArg, HyperlinkColor=defaultNamedOptArg, FollowedHyperlinkColor=defaultNamedOptArg, UnderlineHyperlinks=defaultNamedOptArg, SetDefaults=defaultNamedOptArg
			, ServerIdentification=defaultNamedOptArg, AllowTaskDelegation=defaultNamedOptArg, UpdateProjectToWeb=defaultNamedOptArg, PublishInformationOnSave=defaultNamedOptArg, SetDefaultsMessaging=defaultNamedOptArg
			, SetDefaultsWebServer=defaultNamedOptArg, ManagerEmail=defaultNamedOptArg, ConfirmationDialog=defaultNamedOptArg, ChangesMarkAssnDirty=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2380, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),WorkgroupMessages
			, ServerURL, ServerPath, ReceiveNotifications, SendHyperlinkNote, HyperlinkColor
			, FollowedHyperlinkColor, UnderlineHyperlinks, SetDefaults, ServerIdentification, AllowTaskDelegation
			, UpdateProjectToWeb, PublishInformationOnSave, SetDefaultsMessaging, SetDefaultsWebServer, ManagerEmail
			, ConfirmationDialog, ChangesMarkAssnDirty)

	def Organizer(self, Type=0, Task=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(126, LCID, 1, (11, 0), ((3, 48), (12, 16)),Type
			, Task)

	def OrganizerDeleteItem(self, Type=defaultNamedOptArg, Filename=defaultNamedOptArg, Name=defaultNamedOptArg, Task=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(128, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Type
			, Filename, Name, Task)

	def OrganizerMoveItem(self, Type=defaultNamedOptArg, Filename=defaultNamedOptArg, ToFileName=defaultNamedOptArg, Name=defaultNamedOptArg
			, Task=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(127, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Type
			, Filename, ToFileName, Name, Task)

	def OrganizerRenameItem(self, Type=defaultNamedOptArg, Filename=defaultNamedOptArg, Name=defaultNamedOptArg, NewName=defaultNamedOptArg
			, Task=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(130, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Type
			, Filename, Name, NewName, Task)

	def OutlineHideSubTasks(self):
		return self._oleobj_.InvokeTypes(2020, LCID, 1, (11, 0), (),)

	def OutlineIndent(self, Levels=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2019, LCID, 1, (11, 0), ((12, 16),),Levels
			)

	def OutlineOutdent(self, Levels=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2018, LCID, 1, (11, 0), ((12, 16),),Levels
			)

	def OutlineShowAllTasks(self):
		return self._oleobj_.InvokeTypes(2022, LCID, 1, (11, 0), (),)

	def OutlineShowSubTasks(self):
		return self._oleobj_.InvokeTypes(2021, LCID, 1, (11, 0), (),)

	def OutlineShowTasks(self, OutlineNumber=65535, ExpandInsertedProjects=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(27, LCID, 1, (11, 0), ((3, 48), (12, 16)),OutlineNumber
			, ExpandInsertedProjects)

	def OutlineSymbolsToggle(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2082, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def PERTBorders(self, CriticalStyle=defaultNamedOptArg, CriticalColor=defaultNamedOptArg, NoncriticalStyle=defaultNamedOptArg, NoncriticalColor=defaultNamedOptArg
			, CriticalMilestoneStyle=defaultNamedOptArg, CriticalMilestoneColor=defaultNamedOptArg, NoncriticalMilestoneStyle=defaultNamedOptArg, NoncriticalMilestoneColor=defaultNamedOptArg, CriticalSummaryStyle=defaultNamedOptArg
			, CriticalSummaryColor=defaultNamedOptArg, NoncriticalSummaryStyle=defaultNamedOptArg, NoncriticalSummaryColor=defaultNamedOptArg, CriticalSubprojectStyle=defaultNamedOptArg, CriticalSubprojectColor=defaultNamedOptArg
			, NoncriticalSubprojectStyle=defaultNamedOptArg, NoncriticalSubprojectColor=defaultNamedOptArg, CriticalMarkedStyle=defaultNamedOptArg, CriticalMarkedColor=defaultNamedOptArg, NoncriticalMarkedStyle=defaultNamedOptArg
			, NoncriticalMarkedColor=defaultNamedOptArg, CriticalExternalTaskStyle=defaultNamedOptArg, CriticalExternalTaskColor=defaultNamedOptArg, NoncriticalExternalTaskStyle=defaultNamedOptArg, NoncriticalExternalTaskColor=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(908, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),CriticalStyle
			, CriticalColor, NoncriticalStyle, NoncriticalColor, CriticalMilestoneStyle, CriticalMilestoneColor
			, NoncriticalMilestoneStyle, NoncriticalMilestoneColor, CriticalSummaryStyle, CriticalSummaryColor, NoncriticalSummaryStyle
			, NoncriticalSummaryColor, CriticalSubprojectStyle, CriticalSubprojectColor, NoncriticalSubprojectStyle, NoncriticalSubprojectColor
			, CriticalMarkedStyle, CriticalMarkedColor, NoncriticalMarkedStyle, NoncriticalMarkedColor, CriticalExternalTaskStyle
			, CriticalExternalTaskColor, NoncriticalExternalTaskStyle, NoncriticalExternalTaskColor)

	def PERTBoxStyles(self, Size=defaultNamedOptArg, DateFormat=defaultNamedOptArg, Gridlines=defaultNamedOptArg, CrossMarks=defaultNamedOptArg
			, Field1=defaultNamedOptArg, Field2=defaultNamedOptArg, Field3=defaultNamedOptArg, Field4=defaultNamedOptArg, Field5=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2056, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Size
			, DateFormat, Gridlines, CrossMarks, Field1, Field2
			, Field3, Field4, Field5)

	def PERTLayout(self, Straight=defaultNamedOptArg, DisplayArrows=defaultNamedOptArg, AdjustForPageBreaks=defaultNamedOptArg, DisplayPageBreaks=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(906, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Straight
			, DisplayArrows, AdjustForPageBreaks, DisplayPageBreaks)

	def PERTSetTask(self, Create=defaultNamedOptArg, Move=defaultNamedOptArg, TaskID=defaultNamedOptArg, XPosition=defaultNamedOptArg
			, YPosition=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2055, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Create
			, Move, TaskID, XPosition, YPosition)

	def PageBreakRemove(self):
		return self._oleobj_.InvokeTypes(935, LCID, 1, (11, 0), (),)

	def PageBreakSet(self):
		return self._oleobj_.InvokeTypes(934, LCID, 1, (11, 0), (),)

	def PageBreaksRemoveAll(self):
		return self._oleobj_.InvokeTypes(936, LCID, 1, (11, 0), (),)

	def PageBreaksShow(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(933, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def PanZoomPanTo(self, Start=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5243, LCID, 1, (24, 0), ((12, 0),),Start
			)

	def PanZoomZoomTo(self, Start=defaultNamedNotOptArg, Finish=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5244, LCID, 1, (24, 0), ((12, 0), (12, 0)),Start
			, Finish)

	def PaneClose(self):
		return self._oleobj_.InvokeTypes(2004, LCID, 1, (11, 0), (),)

	def PaneCreate(self):
		return self._oleobj_.InvokeTypes(2003, LCID, 1, (11, 0), (),)

	def PaneNext(self):
		return self._oleobj_.InvokeTypes(2002, LCID, 1, (11, 0), (),)

	def ProgressLines(self):
		return self._oleobj_.InvokeTypes(247, LCID, 1, (11, 0), (),)

	def ProjectMove(self, Date=defaultNamedOptArg, MoveDeadline=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2291, LCID, 1, (11, 0), ((12, 16), (12, 16)),Date
			, MoveDeadline)

	def ProjectStatistics(self, Project=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(602, LCID, 1, (11, 0), ((12, 16),),Project
			)

	def ProjectSummaryInfo(self, Project=defaultNamedOptArg, Title=defaultNamedOptArg, Subject=defaultNamedOptArg, Author=defaultNamedOptArg
			, Company=defaultNamedOptArg, Manager=defaultNamedOptArg, Keywords=defaultNamedOptArg, Comments=defaultNamedOptArg, Start=defaultNamedOptArg
			, Finish=defaultNamedOptArg, ScheduleFrom=defaultNamedOptArg, CurrentDate=defaultNamedOptArg, Calendar=defaultNamedOptArg, StatusDate=defaultNamedOptArg
			, Priority=defaultNamedOptArg, PartiallyDisabled=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(601, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Project
			, Title, Subject, Author, Company, Manager
			, Keywords, Comments, Start, Finish, ScheduleFrom
			, CurrentDate, Calendar, StatusDate, Priority, PartiallyDisabled
			)

	def Publish(self, Republish=defaultNamedOptArg, WssUrl=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2278, LCID, 1, (11, 0), ((12, 16), (12, 16)),Republish
			, WssUrl)

	def PublishAllInformation(self):
		return self._oleobj_.InvokeTypes(5203, LCID, 1, (24, 0), (),)

	def PublishNewAndChangedAssignments(self, ShowDialog=False, ItemsScope=1, NotifyResources=True, NotificationText=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(5201, LCID, 1, (24, 0), ((11, 48), (3, 48), (11, 48), (12, 16)),ShowDialog
			, ItemsScope, NotifyResources, NotificationText)

	def PublishProjectPlan(self, ShowDialog=False, PublishFullPlan=True):
		return self._oleobj_.InvokeTypes(5202, LCID, 1, (24, 0), ((11, 48), (11, 48)),ShowDialog
			, PublishFullPlan)

	def PublisherOptions(self, Name=defaultNamedOptArg, View=defaultNamedOptArg, IsTask=defaultNamedOptArg, UniqueID=defaultNamedOptArg
			, Field=defaultNamedOptArg, OnSave=defaultNamedOptArg, action=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(226, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, View, IsTask, UniqueID, Field, OnSave
			, action)

	def Quit(self, SaveChanges=2):
		return self._oleobj_.InvokeTypes(5008, LCID, 1, (24, 0), ((3, 48),),SaveChanges
			)

	def ReassignSelectedAssns(self, ResourceUniqueID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1512, LCID, 1, (11, 0), ((12, 16),),ResourceUniqueID
			)

	def RecurringTaskInsert(self):
		return self._oleobj_.InvokeTypes(2303, LCID, 1, (11, 0), (),)

	def Redo(self, HowManyRedos=1):
		return self._oleobj_.InvokeTypes(5214, LCID, 1, (11, 0), ((3, 48),),HowManyRedos
			)

	def RegisterProject(self, WaitForCompletion=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5206, LCID, 1, (3, 0), ((11, 0),),WaitForCompletion
			)

	def ReminderSet(self, Start=defaultNamedOptArg, LeadTime=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2383, LCID, 1, (11, 0), ((12, 16), (12, 16)),Start
			, LeadTime)

	def Replace(self, Field=defaultNamedOptArg, Test=defaultNamedOptArg, Value=defaultNamedOptArg, Replacement=defaultNamedOptArg
			, ReplaceAll=defaultNamedOptArg, Next=defaultNamedOptArg, MatchCase=defaultNamedOptArg, FieldID=defaultNamedOptArg, TestID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(241, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Field
			, Test, Value, Replacement, ReplaceAll, Next
			, MatchCase, FieldID, TestID)

	def ReplaceEx(self, Field=defaultNamedOptArg, Test=defaultNamedOptArg, Value=defaultNamedOptArg, Replacement=defaultNamedOptArg
			, ReplaceAll=defaultNamedOptArg, Next=defaultNamedOptArg, MatchCase=defaultNamedOptArg, FieldID=defaultNamedOptArg, TestID=defaultNamedOptArg
			, SearchAllFields=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(98, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Field
			, Test, Value, Replacement, ReplaceAll, Next
			, MatchCase, FieldID, TestID, SearchAllFields)

	def ReplanAssignments(self, action=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2189, LCID, 1, (11, 0), ((12, 16),),action
			)

	def ReportPrint(self, Name=defaultNamedOptArg, FromPage=defaultNamedOptArg, ToPage=defaultNamedOptArg, PageBreaks=defaultNamedOptArg
			, Draft=defaultNamedOptArg, Copies=defaultNamedOptArg, FromDate=defaultNamedOptArg, ToDate=defaultNamedOptArg, Preview=defaultNamedOptArg
			, Color=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(110, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, FromPage, ToPage, PageBreaks, Draft, Copies
			, FromDate, ToDate, Preview, Color)

	def ReportPrintPreview(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(112, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def Reports(self):
		return self._oleobj_.InvokeTypes(2334, LCID, 1, (11, 0), (),)

	def RepublishAssignments(self, ShowDialog=False, ItemsScope=1, NotifyResources=True, OverwriteActuals=False
			, BecomeManager=False, NotificationText=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(5205, LCID, 1, (24, 0), ((11, 48), (3, 48), (11, 48), (11, 48), (11, 48), (12, 16)),ShowDialog
			, ItemsScope, NotifyResources, OverwriteActuals, BecomeManager, NotificationText
			)

	def RequestProgressInformation(self, ShowDialog=False, ItemsScope=1, NotifyTaskLead=False, NotificationText=defaultNamedOptArg
			, ReportingPeriodFrom=defaultNamedOptArg, ReportingPeriodTo=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(5204, LCID, 1, (24, 0), ((11, 48), (3, 48), (11, 48), (12, 16), (12, 16), (12, 16)),ShowDialog
			, ItemsScope, NotifyTaskLead, NotificationText, ReportingPeriodFrom, ReportingPeriodTo
			)

	def RescheduleToNextAvailable(self, TaskID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2193, LCID, 1, (11, 0), ((12, 16),),TaskID
			)

	def ResetTPStyle(self, Style=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1508, LCID, 1, (11, 0), ((3, 0),),Style
			)

	def ResetTrackingMethod(self):
		return self._oleobj_.InvokeTypes(2253, LCID, 1, (11, 0), (),)

	def ResourceActiveDirectory(self):
		return self._oleobj_.InvokeTypes(2200, LCID, 1, (11, 0), (),)

	def ResourceAddressBook(self):
		return self._oleobj_.InvokeTypes(2115, LCID, 1, (11, 0), (),)

	def ResourceAssignment(self, Resources=defaultNamedNotOptArg, Operation=0, With=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(212, LCID, 1, (11, 0), ((12, 16), (3, 48), (12, 16)),Resources
			, Operation, With)

	def ResourceAssignmentDialog(self, ShowResourceListOptions=defaultNamedNotOptArg, ResourceListFields=2, UseNamedFilter=defaultNamedOptArg, FilterName=defaultNamedOptArg
			, UseAvailableToWorkFilter=defaultNamedOptArg, AvailableToWork=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(943, LCID, 1, (11, 0), ((12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16)),ShowResourceListOptions
			, ResourceListFields, UseNamedFilter, FilterName, UseAvailableToWorkFilter, AvailableToWork
			)

	def ResourceCalendarEditDays(self, ProjectName=defaultNamedNotOptArg, ResourceName=defaultNamedNotOptArg, StartDate=defaultNamedOptArg, EndDate=defaultNamedOptArg
			, WeekDay=defaultNamedOptArg, Working=defaultNamedOptArg, Default=defaultNamedOptArg, From1=defaultNamedOptArg, To1=defaultNamedOptArg
			, From2=defaultNamedOptArg, To2=defaultNamedOptArg, From3=defaultNamedOptArg, To3=defaultNamedOptArg, From4=defaultNamedOptArg
			, To4=defaultNamedOptArg, From5=defaultNamedOptArg, To5=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(620, LCID, 1, (11, 0), ((8, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),ProjectName
			, ResourceName, StartDate, EndDate, WeekDay, Working
			, Default, From1, To1, From2, To2
			, From3, To3, From4, To4, From5
			, To5)

	def ResourceCalendarReset(self, ProjectName=defaultNamedNotOptArg, ResourceName=defaultNamedNotOptArg, BaseCalendar=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(621, LCID, 1, (11, 0), ((8, 0), (8, 0), (12, 16)),ProjectName
			, ResourceName, BaseCalendar)

	def ResourceCalendars(self, Index=defaultNamedOptArg, Locked=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(605, LCID, 1, (11, 0), ((12, 16), (12, 16)),Index
			, Locked)

	def ResourceComparison(self):
		return self._oleobj_.InvokeTypes(2185, LCID, 1, (11, 0), (),)

	def ResourceDetails(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2116, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def ResourceGraphBarStyles(self, TopLeftShowAs=defaultNamedOptArg, TopLeftColor=defaultNamedOptArg, TopLeftPattern=defaultNamedOptArg, BottomLeftShowAs=defaultNamedOptArg
			, BottomLeftColor=defaultNamedOptArg, BottomLeftPattern=defaultNamedOptArg, TopRightShowAs=defaultNamedOptArg, TopRightColor=defaultNamedOptArg, TopRightPattern=defaultNamedOptArg
			, BottomRightShowAs=defaultNamedOptArg, BottomRightColor=defaultNamedOptArg, BottomRightPattern=defaultNamedOptArg, ShowValues=defaultNamedOptArg, ShowAvailabilityLine=defaultNamedOptArg
			, PercentBarOverlap=defaultNamedOptArg, ProposedLeftShowAs=defaultNamedOptArg, ProposedLeftColor=defaultNamedOptArg, ProposedLeftPattern=defaultNamedOptArg, ProposedRightShowAs=defaultNamedOptArg
			, ProposedRightColor=defaultNamedOptArg, ProposedRightPattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2057, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),TopLeftShowAs
			, TopLeftColor, TopLeftPattern, BottomLeftShowAs, BottomLeftColor, BottomLeftPattern
			, TopRightShowAs, TopRightColor, TopRightPattern, BottomRightShowAs, BottomRightColor
			, BottomRightPattern, ShowValues, ShowAvailabilityLine, PercentBarOverlap, ProposedLeftShowAs
			, ProposedLeftColor, ProposedLeftPattern, ProposedRightShowAs, ProposedRightColor, ProposedRightPattern
			)

	def ResourceGraphBarStylesEx(self, TopLeftShowAs=defaultNamedOptArg, TopLeftColor=defaultNamedOptArg, TopLeftPattern=defaultNamedOptArg, BottomLeftShowAs=defaultNamedOptArg
			, BottomLeftColor=defaultNamedOptArg, BottomLeftPattern=defaultNamedOptArg, TopRightShowAs=defaultNamedOptArg, TopRightColor=defaultNamedOptArg, TopRightPattern=defaultNamedOptArg
			, BottomRightShowAs=defaultNamedOptArg, BottomRightColor=defaultNamedOptArg, BottomRightPattern=defaultNamedOptArg, ShowValues=defaultNamedOptArg, ShowAvailabilityLine=defaultNamedOptArg
			, PercentBarOverlap=defaultNamedOptArg, ProposedLeftShowAs=defaultNamedOptArg, ProposedLeftColor=defaultNamedOptArg, ProposedLeftPattern=defaultNamedOptArg, ProposedRightShowAs=defaultNamedOptArg
			, ProposedRightColor=defaultNamedOptArg, ProposedRightPattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2153, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),TopLeftShowAs
			, TopLeftColor, TopLeftPattern, BottomLeftShowAs, BottomLeftColor, BottomLeftPattern
			, TopRightShowAs, TopRightColor, TopRightPattern, BottomRightShowAs, BottomRightColor
			, BottomRightPattern, ShowValues, ShowAvailabilityLine, PercentBarOverlap, ProposedLeftShowAs
			, ProposedLeftColor, ProposedLeftPattern, ProposedRightShowAs, ProposedRightColor, ProposedRightPattern
			)

	def ResourceMappingDialog(self):
		return self._oleobj_.InvokeTypes(2255, LCID, 1, (11, 0), (),)

	def ResourceSharing(self, Share=defaultNamedOptArg, Name=defaultNamedOptArg, Pool=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(105, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),Share
			, Name, Pool)

	def ResourceSharingPoolAction(self, action=defaultNamedNotOptArg, Filename=defaultNamedOptArg, ReadOnly=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2083, LCID, 1, (11, 0), ((3, 0), (12, 16), (12, 16)),action
			, Filename, ReadOnly)

	def ResourceSharingPoolRefresh(self):
		return self._oleobj_.InvokeTypes(249, LCID, 1, (11, 0), (),)

	def ResourceSharingPoolUpdate(self, allSharers=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(248, LCID, 1, (11, 0), ((12, 16),),allSharers
			)

	def ResourceWindowsAccount(self, Name=defaultNamedOptArg, ShowDialog=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2394, LCID, 1, (11, 0), ((12, 16), (12, 16)),Name
			, ShowDialog)

	def RestoreSheetSelection(self):
		return self._oleobj_.InvokeTypes(2096, LCID, 1, (11, 0), (),)

	def RowClear(self):
		return self._oleobj_.InvokeTypes(239, LCID, 1, (11, 0), (),)

	def RowDelete(self):
		return self._oleobj_.InvokeTypes(231, LCID, 1, (11, 0), (),)

	def RowInsert(self):
		return self._oleobj_.InvokeTypes(229, LCID, 1, (11, 0), (),)

	def Run(self, Name=defaultNamedNotOptArg, Arg1=defaultNamedOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg
			, Arg4=defaultNamedOptArg, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg
			, Arg9=defaultNamedOptArg, Arg10=defaultNamedOptArg, Arg11=defaultNamedOptArg, Arg12=defaultNamedOptArg, Arg13=defaultNamedOptArg
			, Arg14=defaultNamedOptArg, Arg15=defaultNamedOptArg, Arg16=defaultNamedOptArg, Arg17=defaultNamedOptArg, Arg18=defaultNamedOptArg
			, Arg19=defaultNamedOptArg, Arg20=defaultNamedOptArg, Arg21=defaultNamedOptArg, Arg22=defaultNamedOptArg, Arg23=defaultNamedOptArg
			, Arg24=defaultNamedOptArg, Arg25=defaultNamedOptArg, Arg26=defaultNamedOptArg, Arg27=defaultNamedOptArg, Arg28=defaultNamedOptArg
			, Arg29=defaultNamedOptArg, Arg30=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(5049, LCID, 1, (24, 0), ((8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Arg1, Arg2, Arg3, Arg4, Arg5
			, Arg6, Arg7, Arg8, Arg9, Arg10
			, Arg11, Arg12, Arg13, Arg14, Arg15
			, Arg16, Arg17, Arg18, Arg19, Arg20
			, Arg21, Arg22, Arg23, Arg24, Arg25
			, Arg26, Arg27, Arg28, Arg29, Arg30
			)

	def SaveForSharing(self, Filename=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2133, LCID, 1, (11, 0), ((12, 16),),Filename
			)

	def SaveProjectIfDirty(self, AlertText=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5207, LCID, 1, (11, 0), ((8, 0),),AlertText
			)

	def SaveSheetSelection(self):
		return self._oleobj_.InvokeTypes(2095, LCID, 1, (11, 0), (),)

	def SchedulePlusReminderSet(self, Start=defaultNamedOptArg, LeadTime=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2353, LCID, 1, (11, 0), ((12, 16), (12, 16)),Start
			, LeadTime)

	def SearchFiles(self):
		return self._oleobj_.InvokeTypes(34, LCID, 1, (11, 0), (),)

	def SegmentBorderColor(self, Color=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(72, LCID, 1, (11, 0), ((3, 0),),Color
			)

	def SegmentFillColor(self, Color=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(71, LCID, 1, (11, 0), ((3, 0),),Color
			)

	def SelectAll(self):
		return self._oleobj_.InvokeTypes(216, LCID, 1, (11, 0), (),)

	def SelectBeginning(self, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2041, LCID, 1, (11, 0), ((12, 16),),Extend
			)

	def SelectCell(self, Row=defaultNamedOptArg, Column=defaultNamedOptArg, RowRelative=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2070, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),Row
			, Column, RowRelative)

	def SelectCellDown(self, NumCells=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2050, LCID, 1, (11, 0), ((12, 16), (12, 16)),NumCells
			, Extend)

	def SelectCellLeft(self, NumCells=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2047, LCID, 1, (11, 0), ((12, 16), (12, 16)),NumCells
			, Extend)

	def SelectCellRight(self, NumCells=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2048, LCID, 1, (11, 0), ((12, 16), (12, 16)),NumCells
			, Extend)

	def SelectCellUp(self, NumCells=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2049, LCID, 1, (11, 0), ((12, 16), (12, 16)),NumCells
			, Extend)

	def SelectColumn(self, Column=defaultNamedOptArg, Additional=defaultNamedOptArg, Extend=defaultNamedOptArg, Add=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2046, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Column
			, Additional, Extend, Add)

	def SelectEnd(self, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2042, LCID, 1, (11, 0), ((12, 16),),Extend
			)

	def SelectRange(self, Row=defaultNamedNotOptArg, Column=defaultNamedNotOptArg, RowRelative=defaultNamedOptArg, Width=defaultNamedOptArg
			, Height=defaultNamedOptArg, Extend=defaultNamedOptArg, Add=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2062, LCID, 1, (11, 0), ((3, 0), (2, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Row
			, Column, RowRelative, Width, Height, Extend
			, Add)

	def SelectResourceCell(self, Row=defaultNamedOptArg, Column=defaultNamedOptArg, RowRelative=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2069, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),Row
			, Column, RowRelative)

	def SelectResourceColumn(self, Column=defaultNamedOptArg, Additional=defaultNamedOptArg, Extend=defaultNamedOptArg, Add=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2066, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Column
			, Additional, Extend, Add)

	def SelectResourceField(self, Row=defaultNamedNotOptArg, Column=defaultNamedNotOptArg, RowRelative=defaultNamedOptArg, Width=defaultNamedOptArg
			, Height=defaultNamedOptArg, Extend=defaultNamedOptArg, Add=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2064, LCID, 1, (11, 0), ((3, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Row
			, Column, RowRelative, Width, Height, Extend
			, Add)

	def SelectRow(self, Row=defaultNamedOptArg, RowRelative=defaultNamedOptArg, Height=defaultNamedOptArg, Extend=defaultNamedOptArg
			, Add=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2045, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Row
			, RowRelative, Height, Extend, Add)

	def SelectRowEnd(self, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2044, LCID, 1, (11, 0), ((12, 16),),Extend
			)

	def SelectRowStart(self, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2043, LCID, 1, (11, 0), ((12, 16),),Extend
			)

	def SelectSheet(self):
		return self._oleobj_.InvokeTypes(2067, LCID, 1, (11, 0), (),)

	def SelectTPLineHeight(self, LineMultiple=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(70, LCID, 1, (11, 0), ((2, 0),),LineMultiple
			)

	def SelectTPTask(self, TaskUniqueID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2192, LCID, 1, (11, 0), ((12, 16),),TaskUniqueID
			)

	def SelectTaskAssns(self):
		return self._oleobj_.InvokeTypes(1511, LCID, 1, (11, 0), (),)

	def SelectTaskCell(self, Row=defaultNamedOptArg, Column=defaultNamedOptArg, RowRelative=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2068, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),Row
			, Column, RowRelative)

	def SelectTaskColumn(self, Column=defaultNamedOptArg, Additional=defaultNamedOptArg, Extend=defaultNamedOptArg, Add=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2065, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Column
			, Additional, Extend, Add)

	def SelectTaskField(self, Row=defaultNamedNotOptArg, Column=defaultNamedNotOptArg, RowRelative=defaultNamedOptArg, Width=defaultNamedOptArg
			, Height=defaultNamedOptArg, Extend=defaultNamedOptArg, Add=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2063, LCID, 1, (11, 0), ((3, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Row
			, Column, RowRelative, Width, Height, Extend
			, Add)

	def SelectTimescaleRange(self, Row=defaultNamedNotOptArg, StartTime=defaultNamedNotOptArg, Width=defaultNamedNotOptArg, Height=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(954, LCID, 1, (11, 0), ((3, 0), (8, 0), (2, 0), (3, 0)),Row
			, StartTime, Width, Height)

	def SelectToEnd(self):
		return self._oleobj_.InvokeTypes(1510, LCID, 1, (11, 0), (),)

	def SelectionExtend(self, Extend=defaultNamedOptArg, Add=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2051, LCID, 1, (11, 0), ((12, 16), (12, 16)),Extend
			, Add)

	def ServiceOptionsDialog(self):
		return self._oleobj_.InvokeTypes(2370, LCID, 1, (11, 0), (),)

	def SetActiveCell(self, Value=defaultNamedNotOptArg, Create=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(6, LCID, 1, (11, 0), ((8, 0), (12, 16)),Value
			, Create)

	def SetAutoFilter(self, FieldName=defaultNamedNotOptArg, FilterType=0, Test1=defaultNamedOptArg, Criteria1=defaultNamedOptArg
			, Operation=defaultNamedOptArg, Test2=defaultNamedOptArg, Criteria2=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2166, LCID, 1, (11, 0), ((12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),FieldName
			, FilterType, Test1, Criteria1, Operation, Test2
			, Criteria2)

	def SetField(self, Field=defaultNamedNotOptArg, Value=defaultNamedNotOptArg, Create=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(3, LCID, 1, (11, 0), ((8, 0), (8, 0), (12, 16)),Field
			, Value, Create)

	def SetIgnoreWarningsForTask(self, Set=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2190, LCID, 1, (11, 0), ((12, 16),),Set
			)

	def SetMatchingField(self, Field=defaultNamedNotOptArg, Value=defaultNamedNotOptArg, CheckField=defaultNamedNotOptArg, CheckValue=defaultNamedNotOptArg
			, CheckTest=defaultNamedOptArg, CheckOperation=defaultNamedOptArg, CheckField2=defaultNamedOptArg, CheckValue2=defaultNamedOptArg, CheckTest2=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(11, LCID, 1, (11, 0), ((8, 0), (8, 0), (8, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Field
			, Value, CheckField, CheckValue, CheckTest, CheckOperation
			, CheckField2, CheckValue2, CheckTest2)

	def SetResourceField(self, Field=defaultNamedNotOptArg, Value=defaultNamedNotOptArg, AllSelectedResources=defaultNamedOptArg, Create=defaultNamedOptArg
			, ResourceID=defaultNamedOptArg, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(5, LCID, 1, (11, 0), ((8, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16)),Field
			, Value, AllSelectedResources, Create, ResourceID, ProjectName
			)

	def SetResourceFieldByID(self, FieldID=defaultNamedNotOptArg, Value=defaultNamedNotOptArg, AllSelectedResources=defaultNamedOptArg, Create=defaultNamedOptArg
			, ResourceID=defaultNamedOptArg, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(96, LCID, 1, (11, 0), ((3, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16)),FieldID
			, Value, AllSelectedResources, Create, ResourceID, ProjectName
			)

	def SetRowHeight(self, Unit=defaultNamedOptArg, Rows=defaultNamedOptArg, UseUniqueID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2118, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),Unit
			, Rows, UseUniqueID)

	def SetShowTaskSuggestions(self, Set=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2177, LCID, 1, (11, 0), ((12, 16),),Set
			)

	def SetShowTaskWarnings(self, Set=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2176, LCID, 1, (11, 0), ((12, 16),),Set
			)

	def SetSidepaneStateButton(self, DisplayState=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(5064, LCID, 1, (24, 0), ((11, 0),),DisplayState
			)

	def SetSplitBar(self, ShowColumns=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(31, LCID, 1, (11, 0), ((12, 16),),ShowColumns
			)

	def SetTPField(self, Field=defaultNamedNotOptArg, Value=defaultNamedNotOptArg, AllSelectedTasks=defaultNamedOptArg, Create=defaultNamedOptArg
			, TaskID=defaultNamedOptArg, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1513, LCID, 1, (11, 0), ((8, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16)),Field
			, Value, AllSelectedTasks, Create, TaskID, ProjectName
			)

	def SetTaskField(self, Field=defaultNamedNotOptArg, Value=defaultNamedNotOptArg, AllSelectedTasks=defaultNamedOptArg, Create=defaultNamedOptArg
			, TaskID=defaultNamedOptArg, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(4, LCID, 1, (11, 0), ((8, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16)),Field
			, Value, AllSelectedTasks, Create, TaskID, ProjectName
			)

	def SetTaskFieldByID(self, FieldID=defaultNamedNotOptArg, Value=defaultNamedNotOptArg, AllSelectedTasks=defaultNamedOptArg, Create=defaultNamedOptArg
			, TaskID=defaultNamedOptArg, ProjectName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(95, LCID, 1, (11, 0), ((3, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16)),FieldID
			, Value, AllSelectedTasks, Create, TaskID, ProjectName
			)

	def SetTaskMode(self, Manual=defaultNamedOptArg, IsStickyDates=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(90, LCID, 1, (11, 0), ((12, 16), (12, 16)),Manual
			, IsStickyDates)

	def SetTitleRowHeight(self, TitleHeight=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2120, LCID, 1, (11, 0), ((12, 16),),TitleHeight
			)

	def ShareProjectOnline(self):
		return self._oleobj_.InvokeTypes(74, LCID, 1, (11, 0), (),)

	def ShowAddNewColumn(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(709, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def ShowIgnoredTaskWarnings(self):
		return self._oleobj_.InvokeTypes(2178, LCID, 1, (11, 0), (),)

	def SidepaneTaskChange(self, ID=defaultNamedNotOptArg, IsGoalArea=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(53, LCID, 1, (11, 0), ((2, 0), (12, 16)),ID
			, IsGoalArea)

	def SidepaneToggle(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(52, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def Sort(self, Key1=defaultNamedOptArg, Ascending1=defaultNamedOptArg, Key2=defaultNamedOptArg, Ascending2=defaultNamedOptArg
			, Key3=defaultNamedOptArg, Ascending3=defaultNamedOptArg, Renumber=defaultNamedOptArg, Outline=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(903, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Key1
			, Ascending1, Key2, Ascending2, Key3, Ascending3
			, Renumber, Outline)

	def SpellCheckField(self, FieldName=-1, EnableSpellCheck=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2252, LCID, 1, (11, 0), ((3, 48), (12, 16)),FieldName
			, EnableSpellCheck)

	def SpellingCheck(self):
		return self._oleobj_.InvokeTypes(613, LCID, 1, (11, 0), (),)

	def SplitTask(self, Lock=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1011, LCID, 1, (11, 0), ((12, 16),),Lock
			)

	def StopWebBrowserControlNavigation(self):
		return self._oleobj_.InvokeTypes(55, LCID, 1, (11, 0), (),)

	def SubscribeTo(self, Edition=defaultNamedNotOptArg, Format=1):
		return self._oleobj_.InvokeTypes(223, LCID, 1, (11, 0), ((12, 16), (3, 48)),Edition
			, Format)

	def SubscriberOptions(self, Name=defaultNamedOptArg, IsTask=defaultNamedOptArg, UniqueID=defaultNamedOptArg, Field=defaultNamedOptArg
			, Automatically=defaultNamedOptArg, action=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(225, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, IsTask, UniqueID, Field, Automatically, action
			)

	def SummaryResourceAssignmentsRefresh(self):
		return self._oleobj_.InvokeTypes(2099, LCID, 1, (11, 0), (),)

	def SummaryTasksShow(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(45, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def SynchronizeWithSite(self, SiteURL=defaultNamedOptArg, ListName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2287, LCID, 1, (11, 0), ((12, 16), (12, 16)),SiteURL
			, ListName)

	def TableApply(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(402, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def TableCopy(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(400, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def TableEdit(self, Name=defaultNamedNotOptArg, TaskTable=defaultNamedNotOptArg, Create=defaultNamedOptArg, OverwriteExisting=defaultNamedOptArg
			, NewName=defaultNamedOptArg, FieldName=defaultNamedOptArg, NewFieldName=defaultNamedOptArg, Title=defaultNamedOptArg, Width=defaultNamedOptArg
			, Align=defaultNamedOptArg, ShowInMenu=defaultNamedOptArg, LockFirstColumn=defaultNamedOptArg, DateFormat=defaultNamedOptArg, RowHeight=defaultNamedOptArg
			, ColumnPosition=defaultNamedOptArg, AlignTitle=defaultNamedOptArg, HeaderAutoRowHeightAdjustment=defaultNamedOptArg, HeaderTextWrap=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(403, LCID, 1, (11, 0), ((8, 0), (11, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, TaskTable, Create, OverwriteExisting, NewName, FieldName
			, NewFieldName, Title, Width, Align, ShowInMenu
			, LockFirstColumn, DateFormat, RowHeight, ColumnPosition, AlignTitle
			, HeaderAutoRowHeightAdjustment, HeaderTextWrap)

	def TableEditEx(self, Name=defaultNamedNotOptArg, TaskTable=defaultNamedNotOptArg, Create=defaultNamedOptArg, OverwriteExisting=defaultNamedOptArg
			, NewName=defaultNamedOptArg, FieldName=defaultNamedOptArg, NewFieldName=defaultNamedOptArg, Title=defaultNamedOptArg, Width=defaultNamedOptArg
			, Align=defaultNamedOptArg, ShowInMenu=defaultNamedOptArg, LockFirstColumn=defaultNamedOptArg, DateFormat=defaultNamedOptArg, RowHeight=defaultNamedOptArg
			, ColumnPosition=defaultNamedOptArg, AlignTitle=defaultNamedOptArg, HeaderAutoRowHeightAdjustment=defaultNamedOptArg, HeaderTextWrap=defaultNamedOptArg, WrapText=defaultNamedOptArg
			, ShowAddNewColumn=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2172, LCID, 1, (11, 0), ((8, 0), (11, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, TaskTable, Create, OverwriteExisting, NewName, FieldName
			, NewFieldName, Title, Width, Align, ShowInMenu
			, LockFirstColumn, DateFormat, RowHeight, ColumnPosition, AlignTitle
			, HeaderAutoRowHeightAdjustment, HeaderTextWrap, WrapText, ShowAddNewColumn)

	def TableReset(self):
		return self._oleobj_.InvokeTypes(404, LCID, 1, (11, 0), (),)

	def Tables(self):
		return self._oleobj_.InvokeTypes(401, LCID, 1, (11, 0), (),)

	def TaskComparison(self):
		return self._oleobj_.InvokeTypes(2184, LCID, 1, (11, 0), (),)

	def TaskDeliverableCreate(self, Create=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(92, LCID, 1, (11, 0), ((12, 16),),Create
			)

	def TaskDeliverableSync(self):
		return self._oleobj_.InvokeTypes(93, LCID, 1, (11, 0), (),)

	def TaskDependencySync(self):
		return self._oleobj_.InvokeTypes(94, LCID, 1, (11, 0), (),)

	def TaskDrivers(self):
		return self._oleobj_.InvokeTypes(2279, LCID, 1, (11, 0), (),)

	def TaskInspector(self):
		return self._oleobj_.InvokeTypes(1515, LCID, 1, (11, 0), (),)

	def TaskMove(self, MoveForward=defaultNamedOptArg, IsWorkingDuration=defaultNamedOptArg, MoveDays=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2289, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),MoveForward
			, IsWorkingDuration, MoveDays)

	def TaskMoveToStatusDate(self, MoveCompleted=defaultNamedOptArg, MoveIncomplete=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2290, LCID, 1, (11, 0), ((12, 16), (12, 16)),MoveCompleted
			, MoveIncomplete)

	def TaskOnTimeline(self, TaskID=defaultNamedOptArg, Remove=defaultNamedOptArg, TimelineViewName=defaultNamedOptArg, ShowDialog=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(60, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),TaskID
			, Remove, TimelineViewName, ShowDialog)

	def TaskRespectLinks(self):
		return self._oleobj_.InvokeTypes(2175, LCID, 1, (11, 0), (),)

	def TextStyles(self, Item=defaultNamedOptArg, Font=defaultNamedOptArg, Size=defaultNamedOptArg, Bold=defaultNamedOptArg
			, Italic=defaultNamedOptArg, Underline=defaultNamedOptArg, Color=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(901, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, Font, Size, Bold, Italic, Underline
			, Color)

	def TextStyles32Ex(self, Item=defaultNamedOptArg, Font=defaultNamedOptArg, Size=defaultNamedOptArg, Bold=defaultNamedOptArg
			, Italic=defaultNamedOptArg, Underline=defaultNamedOptArg, Color=defaultNamedOptArg, CellColor=defaultNamedOptArg, Pattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2150, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, Font, Size, Bold, Italic, Underline
			, Color, CellColor, Pattern)

	def TextStylesEx(self, Item=defaultNamedOptArg, Font=defaultNamedOptArg, Size=defaultNamedOptArg, Bold=defaultNamedOptArg
			, Italic=defaultNamedOptArg, Underline=defaultNamedOptArg, Color=defaultNamedOptArg, CellColor=defaultNamedOptArg, Pattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2263, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Item
			, Font, Size, Bold, Italic, Underline
			, Color, CellColor, Pattern)

	def TimelineExport(self, SelectionOnly=defaultNamedOptArg, ExportWidth=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(66, LCID, 1, (11, 0), ((12, 16), (12, 16)),SelectionOnly
			, ExportWidth)

	def TimelineFormat(self, NumLines=defaultNamedOptArg, Minimized=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(64, LCID, 1, (11, 0), ((12, 16), (12, 16)),NumLines
			, Minimized)

	def TimelineGotoSelectedTask(self):
		return self._oleobj_.InvokeTypes(61, LCID, 1, (11, 0), (),)

	def TimelineInsertTask(self, Type=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(65, LCID, 1, (11, 0), ((3, 0),),Type
			)

	def TimelineShowHide(self, Item=defaultNamedNotOptArg, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(62, LCID, 1, (11, 0), ((3, 0), (12, 16)),Item
			, Show)

	def TimelineTextOnBar(self, TextOnBar=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(63, LCID, 1, (11, 0), ((12, 16),),TextOnBar
			)

	def TimelineViewToggle(self):
		return self._oleobj_.InvokeTypes(2194, LCID, 1, (11, 0), (),)

	def Timescale(self):
		return self._oleobj_.InvokeTypes(942, LCID, 1, (11, 0), (),)

	def TimescaleEdit(self, MajorUnits=defaultNamedOptArg, MinorUnits=defaultNamedOptArg, MajorLabel=defaultNamedOptArg, MinorLabel=defaultNamedOptArg
			, MajorAlign=defaultNamedOptArg, MinorAlign=defaultNamedOptArg, MajorCount=defaultNamedOptArg, MinorCount=defaultNamedOptArg, MajorTicks=defaultNamedOptArg
			, MinorTicks=defaultNamedOptArg, Enlarge=defaultNamedOptArg, Separator=defaultNamedOptArg, MajorUseFY=defaultNamedOptArg, MinorUseFY=defaultNamedOptArg
			, TopUnits=defaultNamedOptArg, TopLabel=defaultNamedOptArg, TopAlign=defaultNamedOptArg, TopCount=defaultNamedOptArg, TopTicks=defaultNamedOptArg
			, TopUseFY=defaultNamedOptArg, TierCount=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(902, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),MajorUnits
			, MinorUnits, MajorLabel, MinorLabel, MajorAlign, MinorAlign
			, MajorCount, MinorCount, MajorTicks, MinorTicks, Enlarge
			, Separator, MajorUseFY, MinorUseFY, TopUnits, TopLabel
			, TopAlign, TopCount, TopTicks, TopUseFY, TierCount
			)

	def TimescaleNonWorking(self, Draw=defaultNamedOptArg, Calendar=defaultNamedOptArg, Color=defaultNamedOptArg, Pattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(914, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Draw
			, Calendar, Color, Pattern)

	def TimescaleNonWorkingEx(self, Draw=defaultNamedOptArg, Calendar=defaultNamedOptArg, Color=defaultNamedOptArg, Pattern=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2151, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Draw
			, Calendar, Color, Pattern)

	def TimescaledData(self, TaskID=defaultNamedNotOptArg, ResourceID=defaultNamedNotOptArg, StartDate=defaultNamedNotOptArg, EndDate=defaultNamedNotOptArg
			, Type=0, TimeScaleUnit=3, Count=defaultNamedOptArg):
		return self._ApplyTypes_(4900, 1, (12, 0), ((3, 0), (3, 0), (12, 0), (12, 0), (3, 48), (3, 48), (12, 16)), u'TimescaledData', None,TaskID
			, ResourceID, StartDate, EndDate, Type, TimeScaleUnit
			, Count)

	def TipOfTheDay(self):
		return self._oleobj_.InvokeTypes(2335, LCID, 1, (11, 0), (),)

	def ToggleAssignments(self):
		return self._oleobj_.InvokeTypes(246, LCID, 1, (11, 0), (),)

	def ToggleChangeHighlighting(self):
		return self._oleobj_.InvokeTypes(2136, LCID, 1, (11, 0), (),)

	def TogglePreventResOveralloc(self):
		return self._oleobj_.InvokeTypes(1501, LCID, 1, (11, 0), (),)

	def ToggleResourceDetails(self):
		return self._oleobj_.InvokeTypes(2299, LCID, 1, (11, 0), (),)

	def ToggleTPAutoExpand(self):
		return self._oleobj_.InvokeTypes(1502, LCID, 1, (11, 0), (),)

	def ToggleTPResourceExpand(self, ResourceUniqueID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(73, LCID, 1, (11, 0), ((12, 16),),ResourceUniqueID
			)

	def ToggleTPUnassigned(self):
		return self._oleobj_.InvokeTypes(1506, LCID, 1, (11, 0), (),)

	def ToggleTPUnscheduled(self):
		return self._oleobj_.InvokeTypes(1507, LCID, 1, (11, 0), (),)

	def ToggleTaskDetails(self):
		return self._oleobj_.InvokeTypes(2298, LCID, 1, (11, 0), (),)

	def ToolbarCopyToolFace(self, ToolbarName=defaultNamedNotOptArg, ButtonIndex=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2318, LCID, 1, (11, 0), ((8, 0), (2, 0)),ToolbarName
			, ButtonIndex)

	def ToolbarCustomizeTool(self, ToolbarName=defaultNamedOptArg, ButtonIndex=defaultNamedOptArg, Command=defaultNamedOptArg, FaceIndex=defaultNamedOptArg
			, Description=defaultNamedOptArg, ToolTip=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2316, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),ToolbarName
			, ButtonIndex, Command, FaceIndex, Description, ToolTip
			)

	def ToolbarDeleteTool(self, ToolbarName=defaultNamedNotOptArg, ButtonIndex=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2349, LCID, 1, (11, 0), ((8, 0), (2, 0)),ToolbarName
			, ButtonIndex)

	def ToolbarInsertTool(self, ToolbarName=defaultNamedNotOptArg, ButtonIndex=defaultNamedNotOptArg, Command=defaultNamedOptArg, FaceIndex=defaultNamedOptArg
			, Description=defaultNamedOptArg, ToolTip=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2348, LCID, 1, (11, 0), ((8, 0), (2, 0), (12, 16), (12, 16), (12, 16), (12, 16)),ToolbarName
			, ButtonIndex, Command, FaceIndex, Description, ToolTip
			)

	def ToolbarPasteToolFace(self, ToolbarName=defaultNamedNotOptArg, ButtonIndex=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2319, LCID, 1, (11, 0), ((8, 0), (2, 0)),ToolbarName
			, ButtonIndex)

	def Toolbars(self, action=defaultNamedOptArg, ToolbarName=defaultNamedOptArg, NewToolbarName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2308, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),action
			, ToolbarName, NewToolbarName)

	def ToolbarsCustomize(self):
		return self._oleobj_.InvokeTypes(2309, LCID, 1, (11, 0), (),)

	def Undo(self, HowManyUndos=1):
		return self._oleobj_.InvokeTypes(5213, LCID, 1, (11, 0), ((3, 48),),HowManyUndos
			)

	def UndoClear(self):
		return self._oleobj_.InvokeTypes(5215, LCID, 1, (24, 0), (),)

	def UnlinkTasks(self):
		return self._oleobj_.InvokeTypes(211, LCID, 1, (11, 0), (),)

	def UnloadWebBrowserControl(self, Window=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(5076, LCID, 1, (24, 0), ((16396, 16),),Window
			)

	def UpdateFromProjectServer(self, DataFile=defaultNamedNotOptArg):
		return self._ApplyTypes_(5082, 1, (12, 0), ((8, 0),), u'UpdateFromProjectServer', None,DataFile
			)

	def UpdateProject(self, All=defaultNamedOptArg, UpdateDate=defaultNamedOptArg, action=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(611, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),All
			, UpdateDate, action)

	def UpdateProjectToWeb(self, ServerURL=defaultNamedOptArg, EmbedProjectFile=defaultNamedOptArg, OnlyRegisterProject=defaultNamedOptArg, WaitForCompletion=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2247, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),ServerURL
			, EmbedProjectFile, OnlyRegisterProject, WaitForCompletion)

	def UpdateTasks(self, PercentComplete=defaultNamedOptArg, ActualDuration=defaultNamedOptArg, RemainingDuration=defaultNamedOptArg, ActualStart=defaultNamedOptArg
			, ActualFinish=defaultNamedOptArg, Notes=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2350, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),PercentComplete
			, ActualDuration, RemainingDuration, ActualStart, ActualFinish, Notes
			)

	def UsageViewEntryEx(self, CurIndex=defaultNamedOptArg, Order=defaultNamedOptArg, FontWord=defaultNamedOptArg, CellBackground=defaultNamedOptArg
			, Pattern=defaultNamedOptArg, Shortcut=defaultNamedOptArg, DisplayField=defaultNamedOptArg, FontColor=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2163, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),CurIndex
			, Order, FontWord, CellBackground, Pattern, Shortcut
			, DisplayField, FontColor)

	def ValidateEnterpriseFormula(self, propertyID=defaultNamedNotOptArg, strFormula=defaultNamedNotOptArg, Localized=False):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5217, LCID, 1, (8, 0), ((3, 0), (8, 0), (11, 48)),propertyID
			, strFormula, Localized)

	def ValidateGraphicalIndicators(self, propertyID=defaultNamedNotOptArg, strGraphicalIndicators=defaultNamedNotOptArg, Localized=False):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(5221, LCID, 1, (8, 0), ((3, 0), (8, 0), (11, 48)),propertyID
			, strGraphicalIndicators, Localized)

	def ViewApply(self, Name=defaultNamedOptArg, SinglePane=defaultNamedOptArg, Toggle=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(302, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),Name
			, SinglePane, Toggle)

	def ViewApplyEx(self, Name=defaultNamedNotOptArg, SinglePane=defaultNamedNotOptArg, Toggle=defaultNamedNotOptArg, ApplyTo=defaultNamedNotOptArg
			, BuiltInView=-1):
		return self._oleobj_.InvokeTypes(311, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (3, 48)),Name
			, SinglePane, Toggle, ApplyTo, BuiltInView)

	def ViewBar(self):
		return self._oleobj_.InvokeTypes(966, LCID, 1, (11, 0), (),)

	def ViewCopy(self, Name=defaultNamedOptArg, ApplyTo=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(300, LCID, 1, (11, 0), ((12, 16), (12, 16)),Name
			, ApplyTo)

	def ViewEditCombination(self, Name=defaultNamedOptArg, Create=defaultNamedOptArg, NewName=defaultNamedOptArg, TopView=defaultNamedOptArg
			, BottomView=defaultNamedOptArg, ShowInMenu=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(304, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Create, NewName, TopView, BottomView, ShowInMenu
			)

	def ViewEditSingle(self, Name=defaultNamedOptArg, Create=defaultNamedOptArg, NewName=defaultNamedOptArg, Screen=defaultNamedOptArg
			, ShowInMenu=defaultNamedOptArg, HighlightFilter=defaultNamedOptArg, Table=defaultNamedOptArg, Filter=defaultNamedOptArg, Group=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(303, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Name
			, Create, NewName, Screen, ShowInMenu, HighlightFilter
			, Table, Filter, Group)

	def ViewReset(self):
		return self._oleobj_.InvokeTypes(309, LCID, 1, (11, 0), (),)

	def ViewShowAvailability(self):
		return self._oleobj_.InvokeTypes(927, LCID, 1, (11, 0), (),)

	def ViewShowCost(self):
		return self._oleobj_.InvokeTypes(921, LCID, 1, (11, 0), (),)

	def ViewShowCumulativeCost(self):
		return self._oleobj_.InvokeTypes(928, LCID, 1, (11, 0), (),)

	def ViewShowCumulativeWork(self):
		return self._oleobj_.InvokeTypes(924, LCID, 1, (11, 0), (),)

	def ViewShowNotes(self):
		return self._oleobj_.InvokeTypes(919, LCID, 1, (11, 0), (),)

	def ViewShowObjects(self):
		return self._oleobj_.InvokeTypes(920, LCID, 1, (11, 0), (),)

	def ViewShowOverallocation(self):
		return self._oleobj_.InvokeTypes(925, LCID, 1, (11, 0), (),)

	def ViewShowPeakUnits(self):
		return self._oleobj_.InvokeTypes(923, LCID, 1, (11, 0), (),)

	def ViewShowPercentAllocation(self):
		return self._oleobj_.InvokeTypes(926, LCID, 1, (11, 0), (),)

	def ViewShowPredecessorsSuccessors(self):
		return self._oleobj_.InvokeTypes(917, LCID, 1, (11, 0), (),)

	def ViewShowRemainingAvailability(self):
		return self._oleobj_.InvokeTypes(929, LCID, 1, (11, 0), (),)

	def ViewShowResourcesPredecessors(self):
		return self._oleobj_.InvokeTypes(915, LCID, 1, (11, 0), (),)

	def ViewShowResourcesSuccessors(self):
		return self._oleobj_.InvokeTypes(916, LCID, 1, (11, 0), (),)

	def ViewShowSchedule(self):
		return self._oleobj_.InvokeTypes(918, LCID, 1, (11, 0), (),)

	def ViewShowSelectedTasks(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(932, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def ViewShowUnitAvailability(self):
		return self._oleobj_.InvokeTypes(931, LCID, 1, (11, 0), (),)

	def ViewShowWork(self):
		return self._oleobj_.InvokeTypes(922, LCID, 1, (11, 0), (),)

	def ViewShowWorkAvailability(self):
		return self._oleobj_.InvokeTypes(930, LCID, 1, (11, 0), (),)

	def Views(self):
		return self._oleobj_.InvokeTypes(301, LCID, 1, (11, 0), (),)

	def ViewsEx(self, ApplyTo=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(310, LCID, 1, (11, 0), ((12, 16),),ApplyTo
			)

	def VisualReports(self, PjVisualReportsTab=0):
		return self._oleobj_.InvokeTypes(2137, LCID, 1, (11, 0), ((3, 48),),PjVisualReportsTab
			)

	def VisualReportsEdit(self, strVisualReportTemplateFile=defaultNamedNotOptArg, PjVisualReportsDataLevel=5):
		return self._oleobj_.InvokeTypes(2143, LCID, 1, (11, 0), ((12, 16), (3, 48)),strVisualReportTemplateFile
			, PjVisualReportsDataLevel)

	def VisualReportsNewTemplate(self, PjVisualReportsTemplateType=1, PjVisualReportsCubeType=1, ReportAllFields=defaultNamedNotOptArg, PjVisualReportsDataLevel=5):
		return self._oleobj_.InvokeTypes(2140, LCID, 1, (11, 0), ((3, 48), (3, 48), (12, 16), (3, 48)),PjVisualReportsTemplateType
			, PjVisualReportsCubeType, ReportAllFields, PjVisualReportsDataLevel)

	def VisualReportsSaveCube(self, strNamePath=defaultNamedNotOptArg, PjVisualReportsCubeType=1, ReportAllFields=defaultNamedNotOptArg, PjVisualReportsDataLevel=5):
		return self._oleobj_.InvokeTypes(2139, LCID, 1, (11, 0), ((12, 16), (3, 48), (12, 16), (3, 48)),strNamePath
			, PjVisualReportsCubeType, ReportAllFields, PjVisualReportsDataLevel)

	def VisualReportsSaveDatabase(self, strNamePath=defaultNamedNotOptArg, PjVisualReportsDataLevel=5):
		return self._oleobj_.InvokeTypes(2138, LCID, 1, (11, 0), ((12, 16), (3, 48)),strNamePath
			, PjVisualReportsDataLevel)

	def VisualReportsView(self, strVisualReportTemplateFile=defaultNamedNotOptArg, PjVisualReportsDataLevel=5):
		return self._oleobj_.InvokeTypes(2141, LCID, 1, (11, 0), ((12, 16), (3, 48)),strVisualReportTemplateFile
			, PjVisualReportsDataLevel)

	def WBSCodeMaskEdit(self, CodePrefix=defaultNamedNotOptArg, Level=defaultNamedNotOptArg, Sequence=-1, Length=defaultNamedOptArg
			, Separator=defaultNamedOptArg, CodeGenerate=defaultNamedOptArg, VerifyUniqueness=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(630, LCID, 1, (11, 0), ((12, 16), (12, 16), (3, 48), (12, 16), (12, 16), (12, 16), (12, 16)),CodePrefix
			, Level, Sequence, Length, Separator, CodeGenerate
			, VerifyUniqueness)

	def WBSCodeRenumber(self, All=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(629, LCID, 1, (11, 0), ((12, 16),),All
			)

	def WebAddToFavorites(self, CurrentLink=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1314, LCID, 1, (11, 0), ((12, 16),),CurrentLink
			)

	def WebCopyHyperlink(self):
		return self._oleobj_.InvokeTypes(1313, LCID, 1, (11, 0), (),)

	def WebGoBack(self):
		return self._oleobj_.InvokeTypes(1300, LCID, 1, (11, 0), (),)

	def WebGoForward(self):
		return self._oleobj_.InvokeTypes(1301, LCID, 1, (11, 0), (),)

	def WebHelp(self, Type=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(251, LCID, 1, (11, 0), ((12, 16),),Type
			)

	def WebHideToolbars(self, Hide=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1306, LCID, 1, (11, 0), ((12, 16),),Hide
			)

	def WebInbox(self):
		return self._oleobj_.InvokeTypes(2377, LCID, 1, (11, 0), (),)

	def WebOpenFavorites(self):
		return self._oleobj_.InvokeTypes(1320, LCID, 1, (11, 0), (),)

	def WebOpenHyperlink(self, Address=defaultNamedOptArg, SubAddress=defaultNamedOptArg, AddHistory=defaultNamedOptArg, NewWindow=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1311, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Address
			, SubAddress, AddHistory, NewWindow)

	def WebOpenSearchPage(self):
		return self._oleobj_.InvokeTypes(1305, LCID, 1, (11, 0), (),)

	def WebOpenStartPage(self):
		return self._oleobj_.InvokeTypes(1304, LCID, 1, (11, 0), (),)

	def WebRefresh(self):
		return self._oleobj_.InvokeTypes(1303, LCID, 1, (11, 0), (),)

	def WebSetSearchPage(self, Address=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1318, LCID, 1, (11, 0), ((12, 16),),Address
			)

	def WebSetStartPage(self, Address=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1317, LCID, 1, (11, 0), ((12, 16),),Address
			)

	def WebStopLoading(self):
		return self._oleobj_.InvokeTypes(1302, LCID, 1, (11, 0), (),)

	def WebToolbar(self, Show=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1321, LCID, 1, (11, 0), ((12, 16),),Show
			)

	def WindowActivate(self, WindowName=defaultNamedOptArg, DialogID=defaultNamedOptArg, TopPane=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(705, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),WindowName
			, DialogID, TopPane)

	def WindowArrangeAll(self):
		return self._oleobj_.InvokeTypes(702, LCID, 1, (11, 0), (),)

	def WindowHide(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(703, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def WindowMoreWindows(self):
		return self._oleobj_.InvokeTypes(706, LCID, 1, (11, 0), (),)

	def WindowNewWindow(self, Projects=defaultNamedOptArg, View=defaultNamedOptArg, AllProjects=defaultNamedOptArg, ShowDialog=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(701, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Projects
			, View, AllProjects, ShowDialog)

	def WindowNext(self, NoWrap=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2005, LCID, 1, (11, 0), ((12, 16),),NoWrap
			)

	def WindowPrev(self, NoWrap=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2006, LCID, 1, (11, 0), ((12, 16),),NoWrap
			)

	def WindowSplit(self):
		return self._oleobj_.InvokeTypes(2073, LCID, 1, (11, 0), (),)

	def WindowUnhide(self, Name=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(704, LCID, 1, (11, 0), ((12, 16),),Name
			)

	def WorkOffline(self, fOffline=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2283, LCID, 1, (11, 0), ((12, 16),),fOffline
			)

	def WrapText(self, Column=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(708, LCID, 1, (11, 0), ((12, 16),),Column
			)

	def Zoom(self):
		return self._oleobj_.InvokeTypes(306, LCID, 1, (11, 0), (),)

	def ZoomCalendar(self, NumWeeks=defaultNamedOptArg, StartDate=defaultNamedOptArg, EndDate=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2347, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),NumWeeks
			, StartDate, EndDate)

	def ZoomIn(self):
		return self._oleobj_.InvokeTypes(2036, LCID, 1, (11, 0), (),)

	def ZoomOut(self):
		return self._oleobj_.InvokeTypes(2035, LCID, 1, (11, 0), (),)

	def ZoomTimescale(self, Duration=defaultNamedOptArg, Entire=defaultNamedOptArg, Selection=defaultNamedOptArg, Reset=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(307, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16)),Duration
			, Entire, Selection, Reset)

	_prop_map_get_ = {
		"AMText": (5110, 2, (8, 0), (), "AMText", None),
		# Property 'ActiveCell' is an object of type 'Cell'
		"ActiveCell": (5033, 2, (9, 0), (), "ActiveCell", '{00020B19-0000-0000-C000-000000000046}'),
		# Property 'ActiveProject' is an object of type 'Project'
		"ActiveProject": (5002, 2, (13, 0), (), "ActiveProject", '{1019A320-508A-11CF-A49D-00AA00574C74}'),
		# Property 'ActiveSelection' is an object of type 'Selection'
		"ActiveSelection": (5032, 2, (9, 0), (), "ActiveSelection", '{00020B18-0000-0000-C000-000000000046}'),
		# Property 'ActiveWindow' is an object of type 'Window'
		"ActiveWindow": (5003, 2, (9, 0), (), "ActiveWindow", '{00020B02-0000-0000-C000-000000000046}'),
		# Property 'Application' is an object of type 'Application'
		"Application": (65524, 2, (13, 0), (), "Application", '{36D27C48-A1E8-11D3-BA55-00C04F72F325}'),
		"AskToUpdateLinks": (5155, 2, (11, 0), (), "AskToUpdateLinks", None),
		# Property 'Assistant' is an object of type 'Assistant'
		"Assistant": (5042, 2, (9, 0), (), "Assistant", '{000C0322-0000-0000-C000-000000000046}'),
		"AutoClearLeveling": (5128, 2, (11, 0), (), "AutoClearLeveling", None),
		"AutoLevel": (5127, 2, (11, 0), (), "AutoLevel", None),
		"AutomaticallyFillPhoneticFields": (5165, 2, (11, 0), (), "AutomaticallyFillPhoneticFields", None),
		"AutomationSecurity": (5081, 2, (3, 0), (), "AutomationSecurity", None),
		"Build": (5245, 2, (8, 0), (), "Build", None),
		"Calculation": (5121, 2, (3, 0), (), "Calculation", None),
		"Caption": (5010, 2, (8, 0), (), "Caption", None),
		"CellDragAndDrop": (5139, 2, (11, 0), (), "CellDragAndDrop", None),
		"CompareProjectsCurrentVersionName": (5246, 2, (8, 0), (), "CompareProjectsCurrentVersionName", None),
		"CompareProjectsPreviousVersionName": (5247, 2, (8, 0), (), "CompareProjectsPreviousVersionName", None),
		"CopyResourceUsageHeader": (5138, 2, (11, 0), (), "CopyResourceUsageHeader", None),
		"DateOrder": (5103, 2, (3, 0), (), "DateOrder", None),
		"DateSeparator": (5104, 2, (8, 0), (), "DateSeparator", None),
		"DayLeadingZero": (5158, 2, (11, 0), (), "DayLeadingZero", None),
		"DecimalSeparator": (5106, 2, (8, 0), (), "DecimalSeparator", None),
		"DefaultAutoFilter": (5166, 2, (11, 0), (), "DefaultAutoFilter", None),
		"DefaultDateFormat": (5125, 2, (3, 0), (), "DefaultDateFormat", None),
		"DefaultView": (5112, 2, (8, 0), (), "DefaultView", None),
		"DisplayAlerts": (5157, 2, (11, 0), (), "DisplayAlerts", None),
		"DisplayEntryBar": (5114, 2, (11, 0), (), "DisplayEntryBar", None),
		"DisplayNotesIndicator": (5153, 2, (11, 0), (), "DisplayNotesIndicator", None),
		"DisplayOLEIndicator": (5154, 2, (11, 0), (), "DisplayOLEIndicator", None),
		"DisplayPlanningWizard": (5142, 2, (11, 0), (), "DisplayPlanningWizard", None),
		"DisplayProjectGuide": (5173, 2, (11, 0), (), "DisplayProjectGuide", None),
		"DisplayRecentFiles": (5161, 2, (11, 0), (), "DisplayRecentFiles", None),
		"DisplayScheduleMessages": (5133, 2, (11, 0), (), "DisplayScheduleMessages", None),
		"DisplayScrollBars": (5115, 2, (11, 0), (), "DisplayScrollBars", None),
		"DisplayStatusBar": (5113, 2, (11, 0), (), "DisplayStatusBar", None),
		"DisplayViewBar": (5047, 2, (11, 0), (), "DisplayViewBar", None),
		"DisplayWindowsInTaskbar": (5172, 2, (11, 0), (), "DisplayWindowsInTaskbar", None),
		"DisplayWizardErrors": (5145, 2, (11, 0), (), "DisplayWizardErrors", None),
		"DisplayWizardScheduling": (5143, 2, (11, 0), (), "DisplayWizardScheduling", None),
		"DisplayWizardUsage": (5144, 2, (11, 0), (), "DisplayWizardUsage", None),
		"Edition": (5077, 2, (3, 0), (), "Edition", None),
		"EnableCancelKey": (5044, 2, (3, 0), (), "EnableCancelKey", None),
		"EnableChangeHighlighting": (5222, 2, (11, 0), (), "EnableChangeHighlighting", None),
		"EnterpriseAllowLocalBaseCalendars": (5048, 2, (11, 0), (), "EnterpriseAllowLocalBaseCalendars", None),
		"EnterpriseListSeparator": (5062, 2, (8, 0), (), "EnterpriseListSeparator", None),
		"EnterpriseProtectActuals": (5208, 2, (11, 0), (), "EnterpriseProtectActuals", None),
		# Property 'FileSearch' is an object of type 'FileSearch'
		"FileSearch": (5051, 2, (9, 0), (), "FileSearch", '{000C0332-0000-0000-C000-000000000046}'),
		# Property 'GlobalBaseCalendars' is an object of type 'Calendars'
		"GlobalBaseCalendars": (5083, 2, (9, 0), (), "GlobalBaseCalendars", '{000C0C44-0000-0000-C000-000000000046}'),
		# Property 'GlobalOutlineCodes' is an object of type 'OutlineCodes'
		"GlobalOutlineCodes": (5040, 2, (9, 0), (), "GlobalOutlineCodes", '{4269779B-F035-4E2F-ABDD-54B6D94A2A03}'),
		# Property 'GlobalResourceFilters' is an object of type 'Filters'
		"GlobalResourceFilters": (5074, 2, (9, 0), (), "GlobalResourceFilters", '{ED989E98-F561-4D58-8F67-5D2E2B920E77}'),
		# Property 'GlobalResourceTables' is an object of type 'Tables'
		"GlobalResourceTables": (5072, 2, (9, 0), (), "GlobalResourceTables", '{31E3EB5A-6339-43B0-B1B8-1AED03886AEC}'),
		# Property 'GlobalTaskFilters' is an object of type 'Filters'
		"GlobalTaskFilters": (5073, 2, (9, 0), (), "GlobalTaskFilters", '{ED989E98-F561-4D58-8F67-5D2E2B920E77}'),
		# Property 'GlobalTaskTables' is an object of type 'Tables'
		"GlobalTaskTables": (5071, 2, (9, 0), (), "GlobalTaskTables", '{31E3EB5A-6339-43B0-B1B8-1AED03886AEC}'),
		# Property 'GlobalViews' is an object of type 'Views'
		"GlobalViews": (5068, 2, (9, 0), (), "GlobalViews", '{4CF32482-106B-4FFF-A1AB-7DD395FB0958}'),
		# Property 'GlobalViewsCombination' is an object of type 'ViewsCombination'
		"GlobalViewsCombination": (5070, 2, (9, 0), (), "GlobalViewsCombination", '{CE4F7D83-369B-43CF-96A8-29C2DE2B8658}'),
		# Property 'GlobalViewsSingle' is an object of type 'ViewsSingle'
		"GlobalViewsSingle": (5069, 2, (9, 0), (), "GlobalViewsSingle", '{B4097887-EC12-4683-9622-9974EF26353C}'),
		"Height": (5011, 2, (3, 0), (), "Height", None),
		"Left": (5012, 2, (3, 0), (), "Left", None),
		"LevelFreeformTasks": (5181, 2, (11, 0), (), "LevelFreeformTasks", None),
		"LevelIndividualAssignments": (5168, 2, (11, 0), (), "LevelIndividualAssignments", None),
		"LevelOrder": (5130, 2, (3, 0), (), "LevelOrder", None),
		"LevelPeriodBasis": (5167, 2, (3, 0), (), "LevelPeriodBasis", None),
		"LevelProposedBookings": (5174, 2, (11, 0), (), "LevelProposedBookings", None),
		"LevelWithinSlack": (5129, 2, (11, 0), (), "LevelWithinSlack", None),
		"LevelingCanSplit": (5169, 2, (11, 0), (), "LevelingCanSplit", None),
		"ListSeparator": (5107, 2, (8, 0), (), "ListSeparator", None),
		"LoadLastFile": (5136, 2, (11, 0), (), "LoadLastFile", None),
		"MacroVirusProtection": (5163, 2, (11, 0), (), "MacroVirusProtection", None),
		"MonthLeadingZero": (5159, 2, (11, 0), (), "MonthLeadingZero", None),
		"MoveAfterReturn": (5131, 2, (11, 0), (), "MoveAfterReturn", None),
		"Name": (0, 2, (8, 0), (), "Name", None),
		"NewTasksEstimated": (5171, 2, (11, 0), (), "NewTasksEstimated", None),
		"PMText": (5111, 2, (8, 0), (), "PMText", None),
		"PanZoomFinish": (5242, 2, (12, 0), (), "PanZoomFinish", None),
		"PanZoomStart": (5241, 2, (12, 0), (), "PanZoomStart", None),
		# Property 'Parent' is an object of type 'Application'
		"Parent": (65523, 2, (13, 0), (), "Parent", '{36D27C48-A1E8-11D3-BA55-00C04F72F325}'),
		"Path": (5015, 2, (8, 0), (), "Path", None),
		"PathSeparator": (5007, 2, (8, 0), (), "PathSeparator", None),
		# Property 'Profiles' is an object of type 'Profiles'
		"Profiles": (5038, 2, (9, 0), (), "Profiles", '{08CD6DE7-87CD-42AB-BD28-6E2C0170A274}'),
		# Property 'Projects' is an object of type 'Projects'
		"Projects": (5000, 2, (9, 0), (), "Projects", '{00020B01-0000-0000-C000-000000000046}'),
		"PromptForSummaryInfo": (5149, 2, (11, 0), (), "PromptForSummaryInfo", None),
		"RecentFilesMaximum": (5162, 2, (3, 0), (), "RecentFilesMaximum", None),
		"ScreenUpdating": (5045, 2, (11, 0), (), "ScreenUpdating", None),
		"ShowAssignmentUnitsAs": (5164, 2, (3, 0), (), "ShowAssignmentUnitsAs", None),
		"ShowEstimatedDuration": (5170, 2, (11, 0), (), "ShowEstimatedDuration", None),
		"ShowTipOfDay": (5147, 2, (11, 0), (), "ShowTipOfDay", None),
		"ShowToolTips": (5148, 2, (11, 0), (), "ShowToolTips", None),
		"ShowWelcome": (5156, 2, (11, 0), (), "ShowWelcome", None),
		"StartWeekOn": (5135, 2, (3, 0), (), "StartWeekOn", None),
		"StartYearIn": (5137, 2, (3, 0), (), "StartYearIn", None),
		"StatusBar": (5050, 2, (12, 0), (), "StatusBar", None),
		"SupportsMultipleDocuments": (5004, 2, (11, 0), (), "SupportsMultipleDocuments", None),
		"SupportsMultipleWindows": (5005, 2, (11, 0), (), "SupportsMultipleWindows", None),
		"ThousandSeparator": (5105, 2, (8, 0), (), "ThousandSeparator", None),
		"TimeLeadingZero": (5160, 2, (11, 0), (), "TimeLeadingZero", None),
		"TimeSeparator": (5108, 2, (8, 0), (), "TimeSeparator", None),
		"TimescaleFinish": (5240, 2, (12, 0), (), "TimescaleFinish", None),
		"TimescaleStart": (5239, 2, (12, 0), (), "TimescaleStart", None),
		"Top": (5016, 2, (3, 0), (), "Top", None),
		"TrustProjectServerAndWSSPages": (5180, 2, (11, 0), (), "TrustProjectServerAndWSSPages", None),
		"TwelveHourTimeFormat": (5109, 2, (11, 0), (), "TwelveHourTimeFormat", None),
		"UndoLevels": (5179, 2, (3, 0), (), "UndoLevels", None),
		"UsableHeight": (5017, 2, (5, 0), (), "UsableHeight", None),
		"UsableWidth": (5018, 2, (5, 0), (), "UsableWidth", None),
		"Use3DLook": (5178, 2, (11, 0), (), "Use3DLook", None),
		"UseOMIDs": (5175, 2, (11, 0), (), "UseOMIDs", None),
		"UserControl": (5046, 2, (11, 0), (), "UserControl", None),
		"UserName": (5150, 2, (8, 0), (), "UserName", None),
		# Property 'VBE' is an object of type 'VBE'
		"VBE": (5043, 2, (9, 0), (), "VBE", '{0002E166-0000-0000-C000-000000000046}'),
		"Version": (5020, 2, (8, 0), (), "Version", None),
		"Visible": (5006, 2, (11, 0), (), "Visible", None),
		# Property 'VisualReportTemplateList' is an object of type 'ReportTemplates'
		"VisualReportTemplateList": (5223, 2, (9, 0), (), "VisualReportTemplateList", '{5918F188-19BE-401E-A702-FEE268804738}'),
		"VisualReportsAdditionalTemplatePath": (5231, 2, (8, 0), (), "VisualReportsAdditionalTemplatePath", None),
		"Width": (5021, 2, (3, 0), (), "Width", None),
		"WindowState": (5022, 2, (3, 0), (), "WindowState", None),
		# Method 'AnswerWizard' returns object of type 'AnswerWizard'
		"AnswerWizard": (5056, 2, (9, 0), (), "AnswerWizard", '{000C0360-0000-0000-C000-000000000046}'),
		# Method 'Assistance' returns object of type 'IAssistance'
		"Assistance": (5235, 2, (9, 0), (), "Assistance", '{4291224C-DEFE-485B-8E69-6CF8AA85CB76}'),
		# Method 'COMAddIns' returns object of type 'COMAddIns'
		"COMAddIns": (5055, 2, (9, 0), (), "COMAddIns", '{000C0339-0000-0000-C000-000000000046}'),
		# Method 'CommandBars' returns object of type 'CommandBars'
		"CommandBars": (5041, 2, (13, 0), (), "CommandBars", '{55F88893-7708-11D1-ACEB-006008961DA5}'),
		"OperatingSystem": (5030, 2, (8, 0), (), "OperatingSystem", None),
		# Method 'Windows' returns object of type 'Windows'
		"Windows": (5001, 2, (9, 0), (), "Windows", '{00020B03-0000-0000-C000-000000000046}'),
		# Method 'Windows2' returns object of type 'Windows2'
		"Windows2": (5212, 2, (9, 0), (), "Windows2", '{00020B05-0000-0000-C000-000000000046}'),
	}
	_prop_map_put_ = {
		"AMText" : ((5110, LCID, 4, 0),()),
		"ActiveCell" : ((5033, LCID, 4, 0),()),
		"ActiveProject" : ((5002, LCID, 4, 0),()),
		"ActiveSelection" : ((5032, LCID, 4, 0),()),
		"ActiveWindow" : ((5003, LCID, 4, 0),()),
		"Application" : ((65524, LCID, 4, 0),()),
		"AskToUpdateLinks" : ((5155, LCID, 4, 0),()),
		"Assistant" : ((5042, LCID, 4, 0),()),
		"AutoClearLeveling" : ((5128, LCID, 4, 0),()),
		"AutoLevel" : ((5127, LCID, 4, 0),()),
		"AutomaticallyFillPhoneticFields" : ((5165, LCID, 4, 0),()),
		"AutomationSecurity" : ((5081, LCID, 4, 0),()),
		"Build" : ((5245, LCID, 4, 0),()),
		"Calculation" : ((5121, LCID, 4, 0),()),
		"Caption" : ((5010, LCID, 4, 0),()),
		"CellDragAndDrop" : ((5139, LCID, 4, 0),()),
		"CompareProjectsCurrentVersionName" : ((5246, LCID, 4, 0),()),
		"CompareProjectsPreviousVersionName" : ((5247, LCID, 4, 0),()),
		"CopyResourceUsageHeader" : ((5138, LCID, 4, 0),()),
		"DateOrder" : ((5103, LCID, 4, 0),()),
		"DateSeparator" : ((5104, LCID, 4, 0),()),
		"DayLeadingZero" : ((5158, LCID, 4, 0),()),
		"DecimalSeparator" : ((5106, LCID, 4, 0),()),
		"DefaultAutoFilter" : ((5166, LCID, 4, 0),()),
		"DefaultDateFormat" : ((5125, LCID, 4, 0),()),
		"DefaultView" : ((5112, LCID, 4, 0),()),
		"DisplayAlerts" : ((5157, LCID, 4, 0),()),
		"DisplayEntryBar" : ((5114, LCID, 4, 0),()),
		"DisplayNotesIndicator" : ((5153, LCID, 4, 0),()),
		"DisplayOLEIndicator" : ((5154, LCID, 4, 0),()),
		"DisplayPlanningWizard" : ((5142, LCID, 4, 0),()),
		"DisplayProjectGuide" : ((5173, LCID, 4, 0),()),
		"DisplayRecentFiles" : ((5161, LCID, 4, 0),()),
		"DisplayScheduleMessages" : ((5133, LCID, 4, 0),()),
		"DisplayScrollBars" : ((5115, LCID, 4, 0),()),
		"DisplayStatusBar" : ((5113, LCID, 4, 0),()),
		"DisplayViewBar" : ((5047, LCID, 4, 0),()),
		"DisplayWindowsInTaskbar" : ((5172, LCID, 4, 0),()),
		"DisplayWizardErrors" : ((5145, LCID, 4, 0),()),
		"DisplayWizardScheduling" : ((5143, LCID, 4, 0),()),
		"DisplayWizardUsage" : ((5144, LCID, 4, 0),()),
		"Edition" : ((5077, LCID, 4, 0),()),
		"EnableCancelKey" : ((5044, LCID, 4, 0),()),
		"EnableChangeHighlighting" : ((5222, LCID, 4, 0),()),
		"EnterpriseAllowLocalBaseCalendars" : ((5048, LCID, 4, 0),()),
		"EnterpriseListSeparator" : ((5062, LCID, 4, 0),()),
		"EnterpriseProtectActuals" : ((5208, LCID, 4, 0),()),
		"FileSearch" : ((5051, LCID, 4, 0),()),
		"GlobalBaseCalendars" : ((5083, LCID, 4, 0),()),
		"GlobalOutlineCodes" : ((5040, LCID, 4, 0),()),
		"GlobalResourceFilters" : ((5074, LCID, 4, 0),()),
		"GlobalResourceTables" : ((5072, LCID, 4, 0),()),
		"GlobalTaskFilters" : ((5073, LCID, 4, 0),()),
		"GlobalTaskTables" : ((5071, LCID, 4, 0),()),
		"GlobalViews" : ((5068, LCID, 4, 0),()),
		"GlobalViewsCombination" : ((5070, LCID, 4, 0),()),
		"GlobalViewsSingle" : ((5069, LCID, 4, 0),()),
		"Height" : ((5011, LCID, 4, 0),()),
		"Left" : ((5012, LCID, 4, 0),()),
		"LevelFreeformTasks" : ((5181, LCID, 4, 0),()),
		"LevelIndividualAssignments" : ((5168, LCID, 4, 0),()),
		"LevelOrder" : ((5130, LCID, 4, 0),()),
		"LevelPeriodBasis" : ((5167, LCID, 4, 0),()),
		"LevelProposedBookings" : ((5174, LCID, 4, 0),()),
		"LevelWithinSlack" : ((5129, LCID, 4, 0),()),
		"LevelingCanSplit" : ((5169, LCID, 4, 0),()),
		"ListSeparator" : ((5107, LCID, 4, 0),()),
		"LoadLastFile" : ((5136, LCID, 4, 0),()),
		"MacroVirusProtection" : ((5163, LCID, 4, 0),()),
		"MonthLeadingZero" : ((5159, LCID, 4, 0),()),
		"MoveAfterReturn" : ((5131, LCID, 4, 0),()),
		"Name" : ((0, LCID, 4, 0),()),
		"NewTasksEstimated" : ((5171, LCID, 4, 0),()),
		"PMText" : ((5111, LCID, 4, 0),()),
		"PanZoomFinish" : ((5242, LCID, 4, 0),()),
		"PanZoomStart" : ((5241, LCID, 4, 0),()),
		"Parent" : ((65523, LCID, 4, 0),()),
		"Path" : ((5015, LCID, 4, 0),()),
		"PathSeparator" : ((5007, LCID, 4, 0),()),
		"Profiles" : ((5038, LCID, 4, 0),()),
		"Projects" : ((5000, LCID, 4, 0),()),
		"PromptForSummaryInfo" : ((5149, LCID, 4, 0),()),
		"RecentFilesMaximum" : ((5162, LCID, 4, 0),()),
		"ScreenUpdating" : ((5045, LCID, 4, 0),()),
		"ShowAssignmentUnitsAs" : ((5164, LCID, 4, 0),()),
		"ShowEstimatedDuration" : ((5170, LCID, 4, 0),()),
		"ShowTipOfDay" : ((5147, LCID, 4, 0),()),
		"ShowToolTips" : ((5148, LCID, 4, 0),()),
		"ShowWelcome" : ((5156, LCID, 4, 0),()),
		"StartWeekOn" : ((5135, LCID, 4, 0),()),
		"StartYearIn" : ((5137, LCID, 4, 0),()),
		"StatusBar" : ((5050, LCID, 4, 0),()),
		"SupportsMultipleDocuments" : ((5004, LCID, 4, 0),()),
		"SupportsMultipleWindows" : ((5005, LCID, 4, 0),()),
		"ThousandSeparator" : ((5105, LCID, 4, 0),()),
		"TimeLeadingZero" : ((5160, LCID, 4, 0),()),
		"TimeSeparator" : ((5108, LCID, 4, 0),()),
		"TimescaleFinish" : ((5240, LCID, 4, 0),()),
		"TimescaleStart" : ((5239, LCID, 4, 0),()),
		"Top" : ((5016, LCID, 4, 0),()),
		"TrustProjectServerAndWSSPages" : ((5180, LCID, 4, 0),()),
		"TwelveHourTimeFormat" : ((5109, LCID, 4, 0),()),
		"UndoLevels" : ((5179, LCID, 4, 0),()),
		"UsableHeight" : ((5017, LCID, 4, 0),()),
		"UsableWidth" : ((5018, LCID, 4, 0),()),
		"Use3DLook" : ((5178, LCID, 4, 0),()),
		"UseOMIDs" : ((5175, LCID, 4, 0),()),
		"UserControl" : ((5046, LCID, 4, 0),()),
		"UserName" : ((5150, LCID, 4, 0),()),
		"VBE" : ((5043, LCID, 4, 0),()),
		"Version" : ((5020, LCID, 4, 0),()),
		"Visible" : ((5006, LCID, 4, 0),()),
		"VisualReportTemplateList" : ((5223, LCID, 4, 0),()),
		"VisualReportsAdditionalTemplatePath" : ((5231, LCID, 4, 0),()),
		"Width" : ((5021, LCID, 4, 0),()),
		"WindowState" : ((5022, LCID, 4, 0),()),
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

win32com.client.CLSIDToClass.RegisterCLSID( "{00020AFF-0000-0000-C000-000000000046}", _MSProject )
