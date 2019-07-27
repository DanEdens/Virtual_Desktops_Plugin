# -*- coding: utf-8 -*-
#
# This file is a plugin for EventGhost.
# Copyright © 2005-2019 EventGhost Project <http://www.eventghost.net/>
#
# EventGhost is free software: you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free
# Software Foundation, either version 2 of the License, or (at your option)
# any later version.
#
# EventGhost is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
# FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
# more details.
#
# You should have received a copy of the GNU General Public License along
# with EventGhost. If not, see <http://www.gnu.org/licenses/>.

u"""
    Name: VirtualDesktops
    Author: Kgschlosser
    Version: 0.1
    Description: Creates events based on Virtual desktop interactions.
    GUID: {5DFFBD61-7582-4D6F-8EA9-9CB36284C9CF}
    URL: http://eventghost.net/forum/viewtopic.php?f=10&p=53389#p53389
"""
import eg

eg.RegisterPlugin(
    name = "Virtual Desktops",
    author = "Kgschlosser",
    version = "0.0.4",
    kind="program",
    guid = "{9FF3C497-1518-4036-8A70-969D86F73BE0}",
    canMultiLoad = False,
    url = "http://eventghost.net/forum/viewtopic.php?f=10&p=53389#p53389",
    description = "Creates events based on Virtual desktop interactions.",
    help=(
        u'Will clean up in future revisions\n'
        u'Requires Windows 10 - Build 1903\n'
        u'Events include:\n'
        u'VirtualDesktopCreated\n'
        u'VirtualDesktopDestroyBegin\n'
        u'VirtualDesktopDestroyFailed\n'
        u'VirtualDesktopDestroyed\n'
        u'ViewVirtualDesktopChanged\n'
        u'CurrentVirtualDesktopChanged\n'
        u'IsViewVisible = checks if a desktop is visible or not\n'
        u'IsWindowOnCurrentVirtualDesktop = we will be able to use the Find '
        u'Window action with this one and it will tell us if he window is on '
        u'the current desktop or not.\n'
        u'GetWindowDesktopId = use with Find Windows action to return the id '
        u'of the desktop the window is on\n'
        u'MoveWindowToDesktop = Use with Find Window to move a window to a '
        u'specific desktop\n'
        u'MoveViewToDesktop = I am not 100% sure on this I have to do more '
        u'reading on it.\n'
        u'CanViewMoveDesktops = I am not 100% sure on this I have to do more '
        u'reading on it.\n'
        u'GetCurrentDesktop = Gets the currently selected desktop\n'
        u'GetDesktops = enumerates all desktops\n'
        u'GetAdjacentDesktop = get the desktop either to the left or the right '
        u'of the one that is currently selected\n'
        u'SwitchDesktop = switch to a desktop\n'
        u'CreateDesktop = creates a new desktop\n'
        u'RemoveDesktop = removes a desktop\n'
        u'FindDesktop = finds a desktop'
         )
),

from ctypes.wintypes import HRESULT, HWND, BOOL, POINTER, DWORD, INT, UINT, LPVOID, ULONG

import comtypes
import ctypes
from comtypes import helpstring, COMMETHOD
from comtypes.GUID import GUID


class Text:
    REFGUID = POINTER(GUID)
    REFIID = REFGUID
    ENUM = INT
    IID = GUID
    INT32 = ctypes.c_int32
    INT64 = ctypes.c_int64
    CLSID_ImmersiveShell = GUID(
    '{C2F03A33-21F5-47FA-B4BB-156362A2F239}'
    )
    CLSID_IVirtualNotificationService = GUID(
    '{A501FDEC-4A09-464C-AE4E-1B9C21B84918}'
    )

class VirtualDesktops(eg.PluginBase):
    text = Text

    def __init__(self):
        self.AddAction(IsViewVisible)
        self.AddAction(IsWindowOnCurrentVirtualDesktop)
        self.AddAction(GetWindowDesktopId)
        self.AddAction(MoveWindowToDesktop)
        self.AddAction(MoveViewToDesktop)
        self.AddAction(CanViewMoveDesktops)
        self.AddAction(GetCurrentDesktop)
        self.AddAction(GetDesktops)
        self.AddAction(GetAdjacentDesktop)
        self.AddAction(SwitchDesktop)
        self.AddAction(CreateDesktop)
        self.AddAction(RemoveDesktop)
        self.AddAction(FindDesktop)

    # you will want to add any startup parameters and also run any startup code
    # here
    def __start__(self, *args):
        pass

    # this gets called when eg is being closed and you can run code when that
    # happens
    def __close__(self):
        pass

    # this gets called as well when EG closes but it also gets called when a
    # plugin gets disabled. This is where you will do things like close any
    # open sockets
    def __stop__(self):
        pass

    # You will replace the code in this method if you want to make a plugin
    # configuration dialog.
    def Configure(self, *args):
        eg.PluginBase.Configure(self, *args)

    # The next 2 are pretty self explanatory
    def OnComputerResume(self):
        pass

    def OnComputerSuspend(self):
        pass

    # This gets called when a plugin gets deleted from the tree. so here if
    # you use eg.PersistantData to store any data. that data needs to be
    # deleted when the plugin gets removed. this is where that gets done.
    def OnDelete(self):
        pass


class IsViewVisible(eg.ActionBase):
    name = "IsViewVisible"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class IsWindowOnCurrentVirtualDesktop(eg.ActionBase):
    name = "IsWindowOnCurrentVirtualDesktop"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class GetWindowDesktopId(eg.ActionBase):
    name = "GetWindowDesktopId"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class MoveWindowToDesktop(eg.ActionBase):
    name = "MoveWindowToDesktop"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class MoveViewToDesktop(eg.ActionBase):
    name = "MoveViewToDesktop"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class CanViewMoveDesktops(eg.ActionBase):
    name = "CanViewMoveDesktops"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class GetCurrentDesktop(eg.ActionBase):
    name = "GetCurrentDesktop"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class GetDesktops(eg.ActionBase):
    name = "GetDesktops"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class GetAdjacentDesktop(eg.ActionBase):
    name = "GetAdjacentDesktop"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class SwitchDesktop(eg.ActionBase):
    name = "SwitchDesktop"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class CreateDesktop(eg.ActionBase):
    name = "CreateDesktop"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class RemoveDesktop(eg.ActionBase):
    name = "RemoveDesktop"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class FindDesktop(eg.ActionBase):
    name = "FindDesktop"
    description = ""

    @eg.LogIt
    def __call__(self):
        pass


class HSTRING__(ctypes.Structure):
    _fields_ = [
        ('unused', INT),
    ]


HSTRING = POINTER(HSTRING__)


class EventRegistrationToken(ctypes.Structure):
    _fields_ = [
        ('value', INT64)
    ]


class AdjacentDesktop(ENUM):
    LeftDirection = 3
    RightDirection = 4


class ApplicationViewOrientation(ENUM):
    ApplicationViewOrientation_Landscape = 0
    ApplicationViewOrientation_Portrait = 1


class TrustLevel(ENUM):
    BaseTrust = 0
    PartialTrust = BaseTrust + 1
    FullTrust = PartialTrust + 1


IID_IInspectable = GUID(
    '{AF86E2E0-B12D-4C6A-9C5A-D7AA65101E90}'
)


class IInspectable(comtypes.IUnknown):
    _case_insensitive_ = True
    _idlflags_ = []
    _iid_ = IID_IInspectable
    _methods_ = [
        COMMETHOD(
            [helpstring('Method GetIids')],
            HRESULT,
            'GetIids',
            (['out'], POINTER(ULONG), 'iidCount'),
            (['out'], POINTER(POINTER(IID)), 'iids'),
        ),
        COMMETHOD(
            [helpstring('Method GetRuntimeClassName')],
            HRESULT,
            'GetRuntimeClassName',
            (['out'], POINTER(HSTRING), 'className'),
        ),
        COMMETHOD(
            [helpstring('Method GetTrustLevel')],
            HRESULT,
            'GetTrustLevel',
            (['out'], POINTER(TrustLevel), 'trustLevel'),
        ),
    ]


IID_IApplicationViewConsolidatedEventArgs = GUID(
    '{514449EC-7EA2-4DE7-A6A6-7DFBAAEBB6FB}'
)


class IApplicationViewConsolidatedEventArgs(IInspectable):
    _case_insensitive_ = True
    _iid_ = IID_IApplicationViewConsolidatedEventArgs
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method get_IsUserInitiated')],
            HRESULT,
            'get_IsUserInitiated',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
    ]


IID_IApplicationView = GUID(
    '{D222D519-4361-451E-96C4-60F4F9742DB0}'
)


class IApplicationView(IInspectable):
    _case_insensitive_ = True
    _iid_ = IID_IApplicationView
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method get_Orientation')],
            HRESULT,
            'get_Orientation',
            (['retval', 'out'], POINTER(ApplicationViewOrientation), 'value'),
        ),
        COMMETHOD(
            [helpstring('Method get_AdjacentToLeftDisplayEdge')],
            HRESULT,
            'get_AdjacentToLeftDisplayEdge',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring('Method get_AdjacentToRightDisplayEdge')],
            HRESULT,
            'get_AdjacentToRightDisplayEdge',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring('Method get_IsFullScreen')],
            HRESULT,
            'get_IsFullScreen',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring('Method get_IsOnLockScreen')],
            HRESULT,
            'get_IsOnLockScreen',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring('Method get_IsScreenCaptureEnabled')],
            HRESULT,
            'get_IsScreenCaptureEnabled',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring('Method put_IsScreenCaptureEnabled')],
            HRESULT,
            'put_IsScreenCaptureEnabled',
            (['in'], BOOL, 'value'),
        ),
        COMMETHOD(
            [helpstring('Method put_Title')],
            HRESULT,
            'put_Title',
            (['in'], HSTRING, 'value'),
        ),
        COMMETHOD(
            [helpstring('Method get_Title')],
            HRESULT,
            'get_Title',
            (['retval', 'out'], POINTER(HSTRING), 'value'),
        ),
        COMMETHOD(
            [helpstring('Method get_Id')],
            HRESULT,
            'get_Id',
            (['retval', 'out'], POINTER(INT32), 'value'),
        ),
        COMMETHOD(
            [helpstring('Method add_Consolidated')],
            HRESULT,
            'add_Consolidated',
            (['in'], POINTER(IApplicationViewConsolidatedEventArgs), 'handler'),
            (['retval', 'out'], POINTER(EventRegistrationToken), 'token'),
        ),
        COMMETHOD(
            [helpstring('Method remove_Consolidated')],
            HRESULT,
            'remove_Consolidated',
            (['in', ], EventRegistrationToken, 'EventRegistrationToken'),
        ),
    ]


IID_IServiceProvider = GUID(
    '{6D5140C1-7436-11CE-8034-00AA006009FA}'
)


class IServiceProvider(comtypes.IUnknown):
    _case_insensitive_ = True
    _idlflags_ = []
    _iid_ = IID_IServiceProvider
    _methods_ = [
        COMMETHOD(
            [helpstring('Method QueryService'), 'local', 'in'],
            HRESULT,
            'QueryService',
            (['in'], REFGUID, 'guidService'),
            (['in'], REFIID, 'riid'),
            (['out'], POINTER(LPVOID), 'ppvObject'),
        ),
    ]


IID_IObjectArray = GUID(
    "{92CA9DCD-5622-4BBA-A805-5E9F541BD8C9}"
)


class IObjectArray(comtypes.IUnknown):
    """
    Unknown Object Array
    """
    _case_insensitive_ = True
    _idlflags_ = []
    _iid_ = None

    _methods_ = [
        COMMETHOD(
            [helpstring('Method GetCount')],
            HRESULT,
            'GetCount',
            (['out'], POINTER(UINT), 'pcObjects'),
        ),
        COMMETHOD(
            [helpstring('Method GetAt')],
            HRESULT,
            'GetAt',
            (['in'], UINT, 'uiIndex'),
            (['in'], REFIID, 'riid'),
            (['out', 'iid_is'], POINTER(LPVOID), 'ppv'),
        ),
    ]


IID_IVirtualDesktop = GUID(
    '{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}'
)


class IVirtualDesktop(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktop
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method IsViewVisible')],
            HRESULT,
            'IsViewVisible',
            (['out'], POINTER(IApplicationView), 'pView'),
            (['out'], POINTER(INT), 'pfVisible'),
        ),
        COMMETHOD(
            [helpstring('Method GetID')],
            HRESULT,
            'GetID',
            (['out'], POINTER(GUID), 'pGuid'),
        )
    ]


IID_IVirtualDesktopManager = GUID(
    '{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}'
)


class IVirtualDesktopManager(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktopManager
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method IsWindowOnCurrentVirtualDesktop')],
            HRESULT,
            'IsWindowOnCurrentVirtualDesktop',
            (['in'], HWND, 'topLevelWindow'),
            (['out'], POINTER(BOOL), 'onCurrentDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method GetWindowDesktopId')],
            HRESULT,
            'GetWindowDesktopId',
            (['in'], HWND, 'topLevelWindow'),
            (['out'], POINTER(GUID), 'desktopId'),
        ),
        COMMETHOD(
            [helpstring('Method MoveWindowToDesktop')],
            HRESULT,
            'MoveWindowToDesktop',
            (['in'], HWND, 'topLevelWindow'),
            (['in'], REFGUID, 'desktopId'),
        ),
    ]


CLSID_VirtualDesktopManagerInternal = GUID(
    '{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}'
)

IID_IVirtualDesktopManagerInternal = GUID(
    '{F31574D6-B682-4CDC-BD56-1827860ABEC6}'
)


# IID_IVirtualDesktopManagerInternal = GUID(
#     '{AF8DA486-95BB-4460-B3B7-6E7A6B2962B5}'
# )

# IID_IVirtualDesktopManagerInternal = GUID(
#     '{EF9F1A6C-D3CC-4358-B712-F84B635BEBE7}'
# )

class IVirtualDesktopManagerInternal(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktopManagerInternal
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method GetCount')],
            HRESULT,
            'GetCount',
            (['out'], POINTER(UINT), 'pCount'),
        ),
        COMMETHOD(
            [helpstring('Method MoveViewToDesktop')],
            HRESULT,
            'MoveViewToDesktop',
            (['out'], POINTER(IApplicationView), 'pView'),
            (['out'], POINTER(IVirtualDesktop), 'pDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method CanViewMoveDesktops')],
            HRESULT,
            'CanViewMoveDesktops',
            (['out'], POINTER(IApplicationView), 'pView'),
            (['out'], POINTER(INT), 'pfCanViewMoveDesktops'),
        ),
        COMMETHOD(
            [helpstring('Method GetCurrentDesktop')],
            HRESULT,
            'GetCurrentDesktop',
            (['out'], POINTER(POINTER(IVirtualDesktop)), 'desktop'),
        ),
        COMMETHOD(
            [helpstring('Method GetDesktops')],
            HRESULT,
            'GetDesktops',
            (['out'], POINTER(POINTER(IObjectArray)), 'ppDesktops'),
        ),
        COMMETHOD(
            [helpstring('Method GetAdjacentDesktop')],
            HRESULT,
            'GetAdjacentDesktop',
            (['out'], POINTER(IVirtualDesktop), 'pDesktopReference'),
            (['in'], AdjacentDesktop, 'uDirection'),
            (['out'], POINTER(POINTER(IVirtualDesktop)), 'ppAdjacentDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method SwitchDesktop')],
            HRESULT,
            'SwitchDesktop',
            (['in'], POINTER(IVirtualDesktop), 'pDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method CreateDesktopW')],
            HRESULT,
            'CreateDesktopW',
            (['out'], POINTER(POINTER(IVirtualDesktop)), 'ppNewDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method RemoveDesktop')],
            HRESULT,
            'RemoveDesktop',
            (['in'], POINTER(IVirtualDesktop), 'pRemove'),
            (['in'], POINTER(IVirtualDesktop), 'pFallbackDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method FindDesktop')],
            HRESULT,
            'FindDesktop',
            (['in'], POINTER(GUID), 'desktopId'),
            (['out'], POINTER(POINTER(IVirtualDesktop)), 'ppDesktop'),
        ),
    ]


IID_IVirtualDesktopNotification = GUID(
    '{C179334C-4295-40D3-BEA1-C654D965605A}'
)


class IVirtualDesktopNotification(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktopNotification
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method VirtualDesktopCreated')],
            HRESULT,
            'VirtualDesktopCreated',
            (['in'], POINTER(IVirtualDesktop), 'pDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method VirtualDesktopDestroyBegin')],
            HRESULT,
            'VirtualDesktopDestroyBegin',
            (['in'], POINTER(IVirtualDesktop), 'pDesktopDestroyed'),
            (['in'], POINTER(IVirtualDesktop), 'pDesktopFallback'),
        ),
        COMMETHOD(
            [helpstring('Method VirtualDesktopDestroyFailed')],
            HRESULT,
            'VirtualDesktopDestroyFailed',
            (['in'], POINTER(IVirtualDesktop), 'pDesktopDestroyed'),
            (['in'], POINTER(IVirtualDesktop), 'pDesktopFallback'),
        ),
        COMMETHOD(
            [helpstring('Method VirtualDesktopDestroyed')],
            HRESULT,
            'VirtualDesktopDestroyed',
            (['in'], POINTER(IVirtualDesktop), 'pDesktopDestroyed'),
            (['in'], POINTER(IVirtualDesktop), 'pDesktopFallback'),
        ),
        COMMETHOD(
            [helpstring('Method ViewVirtualDesktopChanged')],
            HRESULT,
            'ViewVirtualDesktopChanged',
            (['in'], POINTER(IApplicationView), 'pView'),
        ),
        COMMETHOD(
            [helpstring('Method CurrentVirtualDesktopChanged')],
            HRESULT,
            'CurrentVirtualDesktopChanged',
            (['in'], POINTER(IVirtualDesktop), 'pDesktopOld'),
            (['in'], POINTER(IVirtualDesktop), 'pDesktopNew'),
        ),
    ]


IID_IVirtualDesktopNotificationService = GUID('{0CD45E71-D927-4F15-8B0A-8FEF525337BF}')


class IVirtualDesktopNotificationService(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktopNotificationService
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method Register')],
            HRESULT,
            'Register',
            (['in'], POINTER(IVirtualDesktopNotification), 'pNotification'),
            (['out'], POINTER(DWORD), 'pdwCookie'),
        ),

        COMMETHOD(
            [helpstring('Method Unregister')],
            HRESULT,
            'Unregister',
            (['in'], DWORD, 'dwCookie'),
        ),
    ]


comtypes.CoInitialize()

pServiceProvider = comtypes.CoCreateInstance(
    CLSID_ImmersiveShell,
    IServiceProvider,
    comtypes.CLSCTX_LOCAL_SERVER,
)

pDesktopManagerInternal = comtypes.cast(
    pServiceProvider.QueryService(
        CLSID_VirtualDesktopManagerInternal,
        IID_IVirtualDesktopManagerInternal
    ),
    ctypes.POINTER(IVirtualDesktopManagerInternal)
)

pObjectArray = POINTER(IObjectArray)()

pDesktopManagerInternal.GetDesktops(ctypes.byref(pObjectArray))

count = UINT()
pObjectArray.GetCount(ctypes.byref(count))

for i in range(count):
    pDesktop = POINTER(IVirtualDesktop)()
    pObjectArray.GetAt(i, IID_IVirtualDesktop, ctypes.byref(pDesktop))

    id = GUID()
    pDesktop.GetID(ctypes.byref(id))

    print(id)
