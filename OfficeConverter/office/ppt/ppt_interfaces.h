#ifndef _PPT_INTERFACES_H_
#define _PPT_INTERFACES_H_

namespace ppt
{

// CApplication 包装器类
class CApplication : public COleDispatchDriver
{
public:
    CApplication(){} // 调用 COleDispatchDriver 默认构造函数
    CApplication(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
    CApplication(const CApplication& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

    // 特性
public:

    // 操作
public:


    // _Application 方法
public:
    LPDISPATCH get_Presentations()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Windows()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPUNKNOWN get_Dialogs()
    {
        LPUNKNOWN result;
        InvokeHelper(0x7d3, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ActiveWindow()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ActivePresentation()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SlideShowWindows()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_CommandBars()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_Path()
    {
        CString result;
        InvokeHelper(0x7d8, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_Name()
    {
        CString result;
        InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_Caption()
    {
        CString result;
        InvokeHelper(0x7d9, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_Caption(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7d9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Assistant()
    {
        LPDISPATCH result;
        InvokeHelper(0x7da, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FileSearch()
    {
        LPDISPATCH result;
        InvokeHelper(0x7db, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FileFind()
    {
        LPDISPATCH result;
        InvokeHelper(0x7dc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_Build()
    {
        CString result;
        InvokeHelper(0x7dd, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_Version()
    {
        CString result;
        InvokeHelper(0x7de, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_OperatingSystem()
    {
        CString result;
        InvokeHelper(0x7df, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_ActivePrinter()
    {
        CString result;
        InvokeHelper(0x7e0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    long get_Creator()
    {
        long result;
        InvokeHelper(0x7e1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_AddIns()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_VBE()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void Help(LPCTSTR HelpFile, long ContextID)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4;
        InvokeHelper(0x7e4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, HelpFile, ContextID);
    }
    void Quit()
    {
        InvokeHelper(0x7e5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    VARIANT Run(LPCTSTR MacroName, SAFEARRAY * * safeArrayOfParams)
    {
        VARIANT result;
        static BYTE parms[] = VTS_BSTR VTS_UNKNOWN;
        InvokeHelper(0x7e6, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, MacroName, safeArrayOfParams);
        return result;
    }
    LPUNKNOWN PPFileDialog(long Type)
    {
        LPUNKNOWN result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7e7, DISPATCH_METHOD, VT_UNKNOWN, (void*)&result, parms, Type);
        return result;
    }
    void LaunchSpelling(LPDISPATCH pWindow)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x7e8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, pWindow);
    }
    float get_Left()
    {
        float result;
        InvokeHelper(0x7e9, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_Left(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x7e9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_Top()
    {
        float result;
        InvokeHelper(0x7ea, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_Top(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x7ea, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_Width()
    {
        float result;
        InvokeHelper(0x7eb, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_Width(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x7eb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_Height()
    {
        float result;
        InvokeHelper(0x7ec, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_Height(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x7ec, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_WindowState()
    {
        long result;
        InvokeHelper(0x7ed, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_WindowState(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7ed, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Visible()
    {
        long result;
        InvokeHelper(0x7ee, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Visible(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7ee, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_HWND()
    {
        long result;
        InvokeHelper(0x7ef, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_Active()
    {
        long result;
        InvokeHelper(0x7f0, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void Activate()
    {
        InvokeHelper(0x7f1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_AnswerWizard()
    {
        LPDISPATCH result;
        InvokeHelper(0x7f2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_COMAddIns()
    {
        LPDISPATCH result;
        InvokeHelper(0x7f3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_ProductCode()
    {
        CString result;
        InvokeHelper(0x7f4, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_DefaultWebOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x7f5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_LanguageSettings()
    {
        LPDISPATCH result;
        InvokeHelper(0x7f6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_MsoDebugOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x7f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_ShowWindowsInTaskbar()
    {
        long result;
        InvokeHelper(0x7f8, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_ShowWindowsInTaskbar(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7f8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPUNKNOWN get_Marker()
    {
        LPUNKNOWN result;
        InvokeHelper(0x7f9, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
        return result;
    }
    long get_FeatureInstall()
    {
        long result;
        InvokeHelper(0x7fa, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_FeatureInstall(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7fa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL GetOptionFlag(long Option, BOOL Persist)
    {
        BOOL result;
        static BYTE parms[] = VTS_I4 VTS_BOOL;
        InvokeHelper(0x7fb, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Option, Persist);
        return result;
    }
    void SetOptionFlag(long Option, BOOL State, BOOL Persist)
    {
        static BYTE parms[] = VTS_I4 VTS_BOOL VTS_BOOL;
        InvokeHelper(0x7fc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Option, State, Persist);
    }
    LPDISPATCH get_FileDialog(long Type)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7fd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, Type);
        return result;
    }
    long get_DisplayGridLines()
    {
        long result;
        InvokeHelper(0x7fe, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_DisplayGridLines(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7fe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_AutomationSecurity()
    {
        long result;
        InvokeHelper(0x7ff, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_AutomationSecurity(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7ff, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_NewPresentation()
    {
        LPDISPATCH result;
        InvokeHelper(0x800, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_DisplayAlerts()
    {
        long result;
        InvokeHelper(0x801, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_DisplayAlerts(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x801, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_ShowStartupDialog()
    {
        long result;
        InvokeHelper(0x802, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_ShowStartupDialog(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x802, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void SetPerfMarker(long Marker)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x803, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Marker);
    }
    LPDISPATCH get_AutoCorrect()
    {
        LPDISPATCH result;
        InvokeHelper(0x804, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Options()
    {
        LPDISPATCH result;
        InvokeHelper(0x805, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void LaunchPublishSlidesDialog(LPCTSTR SlideLibraryUrl)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x806, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SlideLibraryUrl);
    }
    void LaunchSendToPPTDialog(VARIANT * SlideUrls)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x807, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SlideUrls);
    }
    BOOL get_DisplayDocumentInformationPanel()
    {
        BOOL result;
        InvokeHelper(0x808, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DisplayDocumentInformationPanel(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x808, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Assistance()
    {
        LPDISPATCH result;
        InvokeHelper(0x809, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_ActiveEncryptionSession()
    {
        long result;
        InvokeHelper(0x80a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FileConverters()
    {
        LPDISPATCH result;
        InvokeHelper(0x80b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SmartArtLayouts()
    {
        LPDISPATCH result;
        InvokeHelper(0x80c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SmartArtQuickStyles()
    {
        LPDISPATCH result;
        InvokeHelper(0x80d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SmartArtColors()
    {
        LPDISPATCH result;
        InvokeHelper(0x80e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ProtectedViewWindows()
    {
        LPDISPATCH result;
        InvokeHelper(0x80f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ActiveProtectedViewWindow()
    {
        LPDISPATCH result;
        InvokeHelper(0x810, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_IsSandboxed()
    {
        BOOL result;
        InvokeHelper(0x811, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ResampleMediaTasks()
    {
        LPDISPATCH result;
        InvokeHelper(0x812, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void StartNewUndoEntry()
    {
        InvokeHelper(0x813, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    long get_FileValidation()
    {
        long result;
        InvokeHelper(0x814, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_FileValidation(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x814, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH OpenThemeFile(LPCTSTR themeFileName)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x815, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, themeFileName);
        return result;
    }
    BOOL get_ChartDataPointTrack()
    {
        BOOL result;
        InvokeHelper(0x816, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ChartDataPointTrack(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x816, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_DisplayGuides()
    {
        long result;
        InvokeHelper(0x817, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_DisplayGuides(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x817, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }

    // _Application 属性
public:

};

// CPresentations 包装器类
class CPresentations : public COleDispatchDriver
{
public:
    CPresentations(){} // 调用 COleDispatchDriver 默认构造函数
    CPresentations(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
    CPresentations(const CPresentations& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

    // 特性
public:

    // 操作
public:


    // Presentations 方法
public:
    LPUNKNOWN get__NewEnum()
    {
        LPUNKNOWN result;
        InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
        return result;
    }
    VARIANT _Index(long Index)
    {
        VARIANT result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0xa, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, Index);
        return result;
    }
    long get_Count()
    {
        long result;
        InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Application()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Parent()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH Item(VARIANT& Index)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_VARIANT;
        InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
        return result;
    }
    LPDISPATCH Add(long WithWindow)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7d3, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, WithWindow);
        return result;
    }
    LPDISPATCH OpenOld(LPCTSTR FileName, long ReadOnly, long Untitled, long WithWindow)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4 VTS_I4;
        InvokeHelper(0x7d4, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, ReadOnly, Untitled, WithWindow);
        return result;
    }
    LPDISPATCH Open(LPCTSTR FileName, long ReadOnly, long Untitled, long WithWindow)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4 VTS_I4;
        InvokeHelper(0x7d5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, ReadOnly, Untitled, WithWindow);
        return result;
    }
    void CheckOut(LPCTSTR FileName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7d6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
    }
    BOOL CanCheckOut(LPCTSTR FileName)
    {
        BOOL result;
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7d7, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, FileName);
        return result;
    }
    LPDISPATCH Open2007(LPCTSTR FileName, long ReadOnly, long Untitled, long WithWindow, long OpenAndRepair)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4 VTS_I4 VTS_I4;
        InvokeHelper(0x7d8, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair);
        return result;
    }

    // Presentations 属性
public:

};

// CPresentation 包装器类
class CPresentation : public COleDispatchDriver
{
public:
    CPresentation(){} // 调用 COleDispatchDriver 默认构造函数
    CPresentation(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
    CPresentation(const CPresentation& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

    // 特性
public:

    // 操作
public:


    // _Presentation 方法
public:
    LPDISPATCH get_Application()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Parent()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SlideMaster()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_TitleMaster()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_HasTitleMaster()
    {
        long result;
        InvokeHelper(0x7d5, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH AddTitleMaster()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d6, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ApplyTemplate(LPCTSTR FileName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7d7, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
    }
    CString get_TemplateName()
    {
        CString result;
        InvokeHelper(0x7d8, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_NotesMaster()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_HandoutMaster()
    {
        LPDISPATCH result;
        InvokeHelper(0x7da, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Slides()
    {
        LPDISPATCH result;
        InvokeHelper(0x7db, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_PageSetup()
    {
        LPDISPATCH result;
        InvokeHelper(0x7dc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ColorSchemes()
    {
        LPDISPATCH result;
        InvokeHelper(0x7dd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ExtraColors()
    {
        LPDISPATCH result;
        InvokeHelper(0x7de, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SlideShowSettings()
    {
        LPDISPATCH result;
        InvokeHelper(0x7df, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Fonts()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Windows()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Tags()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_DefaultShape()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_BuiltInDocumentProperties()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_CustomDocumentProperties()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_VBProject()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_ReadOnly()
    {
        long result;
        InvokeHelper(0x7e7, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    CString get_FullName()
    {
        CString result;
        InvokeHelper(0x7e8, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_Name()
    {
        CString result;
        InvokeHelper(0x7e9, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_Path()
    {
        CString result;
        InvokeHelper(0x7ea, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    long get_Saved()
    {
        long result;
        InvokeHelper(0x7eb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Saved(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7eb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_LayoutDirection()
    {
        long result;
        InvokeHelper(0x7ec, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_LayoutDirection(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7ec, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH NewWindow()
    {
        LPDISPATCH result;
        InvokeHelper(0x7ed, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void FollowHyperlink(LPCTSTR Address, LPCTSTR SubAddress, BOOL NewWindow, BOOL AddHistory, LPCTSTR ExtraInfo, long Method, LPCTSTR HeaderInfo)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BOOL VTS_BOOL VTS_BSTR VTS_I4 VTS_BSTR;
        InvokeHelper(0x7ee, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo);
    }
    void AddToFavorites()
    {
        InvokeHelper(0x7ef, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Unused()
    {
        InvokeHelper(0x7f0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_PrintOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x7f1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void PrintOut(long From, long To, LPCTSTR PrintToFile, long Copies, long Collate)
    {
        static BYTE parms[] = VTS_I4 VTS_I4 VTS_BSTR VTS_I4 VTS_I4;
        InvokeHelper(0x7f2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, From, To, PrintToFile, Copies, Collate);
    }
    void Save()
    {
        InvokeHelper(0x7f3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SaveAs(LPCTSTR FileName, long FileFormat, long EmbedTrueTypeFonts)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4;
        InvokeHelper(0x7f4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, FileFormat, EmbedTrueTypeFonts);
    }
    void SaveCopyAs(LPCTSTR FileName, long FileFormat, long EmbedTrueTypeFonts)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4;
        InvokeHelper(0x7f5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, FileFormat, EmbedTrueTypeFonts);
    }
    void Export(LPCTSTR Path, LPCTSTR FilterName, long ScaleWidth, long ScaleHeight)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_I4 VTS_I4;
        InvokeHelper(0x7f6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Path, FilterName, ScaleWidth, ScaleHeight);
    }
    void Close()
    {
        InvokeHelper(0x7f7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SetUndoText(LPCTSTR Text)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7f8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text);
    }
    LPDISPATCH get_Container()
    {
        LPDISPATCH result;
        InvokeHelper(0x7f9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_DisplayComments()
    {
        long result;
        InvokeHelper(0x7fa, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_DisplayComments(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7fa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_FarEastLineBreakLevel()
    {
        long result;
        InvokeHelper(0x7fb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_FarEastLineBreakLevel(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7fb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_NoLineBreakBefore()
    {
        CString result;
        InvokeHelper(0x7fc, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_NoLineBreakBefore(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7fc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_NoLineBreakAfter()
    {
        CString result;
        InvokeHelper(0x7fd, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_NoLineBreakAfter(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7fd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void UpdateLinks()
    {
        InvokeHelper(0x7fe, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_SlideShowWindow()
    {
        LPDISPATCH result;
        InvokeHelper(0x7ff, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_FarEastLineBreakLanguage()
    {
        long result;
        InvokeHelper(0x800, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_FarEastLineBreakLanguage(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x800, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void WebPagePreview()
    {
        InvokeHelper(0x801, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    long get_DefaultLanguageID()
    {
        long result;
        InvokeHelper(0x802, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_DefaultLanguageID(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x802, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_CommandBars()
    {
        LPDISPATCH result;
        InvokeHelper(0x803, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_PublishObjects()
    {
        LPDISPATCH result;
        InvokeHelper(0x804, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_WebOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x805, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_HTMLProject()
    {
        LPDISPATCH result;
        InvokeHelper(0x806, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ReloadAs(long cp)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x807, DISPATCH_METHOD, VT_EMPTY, NULL, parms, cp);
    }
    void MakeIntoTemplate(long IsDesignTemplate)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x808, DISPATCH_METHOD, VT_EMPTY, NULL, parms, IsDesignTemplate);
    }
    long get_EnvelopeVisible()
    {
        long result;
        InvokeHelper(0x809, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_EnvelopeVisible(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x809, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void sblt(LPCTSTR s)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x80a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, s);
    }
    long get_VBASigned()
    {
        long result;
        InvokeHelper(0x80b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_SnapToGrid()
    {
        long result;
        InvokeHelper(0x80d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_SnapToGrid(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x80d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_GridDistance()
    {
        float result;
        InvokeHelper(0x80e, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_GridDistance(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x80e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Designs()
    {
        LPDISPATCH result;
        InvokeHelper(0x80f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void Merge(LPCTSTR Path)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x810, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Path);
    }
    void CheckIn(BOOL SaveChanges, VARIANT& Comments, VARIANT& MakePublic)
    {
        static BYTE parms[] = VTS_BOOL VTS_VARIANT VTS_VARIANT;
        InvokeHelper(0x811, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SaveChanges, &Comments, &MakePublic);
    }
    BOOL CanCheckIn()
    {
        BOOL result;
        InvokeHelper(0x812, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Signatures()
    {
        LPDISPATCH result;
        InvokeHelper(0x813, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_RemovePersonalInformation()
    {
        long result;
        InvokeHelper(0x814, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_RemovePersonalInformation(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x814, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void SendForReview(LPCTSTR Recipients, LPCTSTR Subject, BOOL ShowMessage, VARIANT& IncludeAttachment)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BOOL VTS_VARIANT;
        InvokeHelper(0x815, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Recipients, Subject, ShowMessage, &IncludeAttachment);
    }
    void ReplyWithChanges(BOOL ShowMessage)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x816, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ShowMessage);
    }
    void EndReview()
    {
        InvokeHelper(0x817, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    long get_HasRevisionInfo()
    {
        long result;
        InvokeHelper(0x818, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void AddBaseline(LPCTSTR FileName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x819, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
    }
    void RemoveBaseline()
    {
        InvokeHelper(0x81a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    CString get_PasswordEncryptionProvider()
    {
        CString result;
        InvokeHelper(0x81b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_PasswordEncryptionAlgorithm()
    {
        CString result;
        InvokeHelper(0x81c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    long get_PasswordEncryptionKeyLength()
    {
        long result;
        InvokeHelper(0x81d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    BOOL get_PasswordEncryptionFileProperties()
    {
        BOOL result;
        InvokeHelper(0x81e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void SetPasswordEncryptionOptions(LPCTSTR PasswordEncryptionProvider, LPCTSTR PasswordEncryptionAlgorithm, long PasswordEncryptionKeyLength, BOOL PasswordEncryptionFileProperties)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_I4 VTS_BOOL;
        InvokeHelper(0x81f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties);
    }
    CString get_Password()
    {
        CString result;
        InvokeHelper(0x820, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_Password(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x820, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_WritePassword()
    {
        CString result;
        InvokeHelper(0x821, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_WritePassword(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x821, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Permission()
    {
        LPDISPATCH result;
        InvokeHelper(0x822, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SharedWorkspace()
    {
        LPDISPATCH result;
        InvokeHelper(0x823, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Sync()
    {
        LPDISPATCH result;
        InvokeHelper(0x824, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void SendFaxOverInternet(LPCTSTR Recipients, LPCTSTR Subject, BOOL ShowMessage)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BOOL;
        InvokeHelper(0x825, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Recipients, Subject, ShowMessage);
    }
    LPDISPATCH get_DocumentLibraryVersions()
    {
        LPDISPATCH result;
        InvokeHelper(0x826, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ContentTypeProperties()
    {
        LPDISPATCH result;
        InvokeHelper(0x827, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_SectionCount()
    {
        long result;
        InvokeHelper(0x828, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    BOOL get_HasSections()
    {
        BOOL result;
        InvokeHelper(0x829, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void NewSectionAfter(long Index, BOOL AfterSlide, LPCTSTR sectionTitle, long * newSectionIndex)
    {
        static BYTE parms[] = VTS_I4 VTS_BOOL VTS_BSTR VTS_PI4;
        InvokeHelper(0x82a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index, AfterSlide, sectionTitle, newSectionIndex);
    }
    void DeleteSection(long Index)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x82b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index);
    }
    void DisableSections()
    {
        InvokeHelper(0x82c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    CString sectionTitle(long Index)
    {
        CString result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x82d, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Index);
        return result;
    }
    void RemoveDocumentInformation(long Type)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x82e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type);
    }
    void CheckInWithVersion(BOOL SaveChanges, VARIANT& Comments, VARIANT& MakePublic, VARIANT& VersionType)
    {
        static BYTE parms[] = VTS_BOOL VTS_VARIANT VTS_VARIANT VTS_VARIANT;
        InvokeHelper(0x82f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SaveChanges, &Comments, &MakePublic, &VersionType);
    }
    void ExportAsFixedFormat(LPCTSTR Path, long FixedFormatType, long Intent, long FrameSlides, long HandoutOrder, long OutputType, long PrintHiddenSlides, LPDISPATCH PrintRange, long RangeType, LPCTSTR SlideShowName, BOOL IncludeDocProperties, BOOL KeepIRMSettings, BOOL DocStructureTags, BOOL BitmapMissingFonts, BOOL UseISO19005_1, VARIANT& ExternalExporter)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_DISPATCH VTS_I4 VTS_BSTR VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_VARIANT;
        InvokeHelper(0x830, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Path, FixedFormatType, Intent, FrameSlides, HandoutOrder, OutputType, PrintHiddenSlides, PrintRange, RangeType, SlideShowName, IncludeDocProperties, KeepIRMSettings, DocStructureTags, BitmapMissingFonts, UseISO19005_1, &ExternalExporter);
    }
    LPDISPATCH get_ServerPolicy()
    {
        LPDISPATCH result;
        InvokeHelper(0x831, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH GetWorkflowTasks()
    {
        LPDISPATCH result;
        InvokeHelper(0x832, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH GetWorkflowTemplates()
    {
        LPDISPATCH result;
        InvokeHelper(0x833, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void LockServerFile()
    {
        InvokeHelper(0x834, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_DocumentInspectors()
    {
        LPDISPATCH result;
        InvokeHelper(0x835, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_HasVBProject()
    {
        BOOL result;
        InvokeHelper(0x836, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_CustomXMLParts()
    {
        LPDISPATCH result;
        InvokeHelper(0x837, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_Final()
    {
        BOOL result;
        InvokeHelper(0x838, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_Final(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x838, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void ApplyTheme(LPCTSTR themeName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x839, DISPATCH_METHOD, VT_EMPTY, NULL, parms, themeName);
    }
    LPDISPATCH get_CustomerData()
    {
        LPDISPATCH result;
        InvokeHelper(0x83a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Research()
    {
        LPDISPATCH result;
        InvokeHelper(0x83b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void PublishSlides(LPCTSTR SlideLibraryUrl, BOOL Overwrite, BOOL UseSlideOrder)
    {
        static BYTE parms[] = VTS_BSTR VTS_BOOL VTS_BOOL;
        InvokeHelper(0x83c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SlideLibraryUrl, Overwrite, UseSlideOrder);
    }
    CString get_EncryptionProvider()
    {
        CString result;
        InvokeHelper(0x83d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_EncryptionProvider(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x83d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void Convert()
    {
        InvokeHelper(0x83e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_SectionProperties()
    {
        LPDISPATCH result;
        InvokeHelper(0x83f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Coauthoring()
    {
        LPDISPATCH result;
        InvokeHelper(0x840, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void MergeWithBaseline(LPCTSTR withPresentation, LPCTSTR baselinePresentation)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR;
        InvokeHelper(0x841, DISPATCH_METHOD, VT_EMPTY, NULL, parms, withPresentation, baselinePresentation);
    }
    BOOL get_InMergeMode()
    {
        BOOL result;
        InvokeHelper(0x842, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void AcceptAll()
    {
        InvokeHelper(0x843, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void RejectAll()
    {
        InvokeHelper(0x844, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void EnsureAllMediaUpgraded()
    {
        InvokeHelper(0x845, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_Broadcast()
    {
        LPDISPATCH result;
        InvokeHelper(0x846, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_HasNotesMaster()
    {
        BOOL result;
        InvokeHelper(0x847, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    BOOL get_HasHandoutMaster()
    {
        BOOL result;
        InvokeHelper(0x848, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void Convert2(LPCTSTR FileName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x849, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
    }
    long get_CreateVideoStatus()
    {
        long result;
        InvokeHelper(0x84a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void CreateVideo(LPCTSTR FileName, BOOL UseTimingsAndNarrations, long DefaultSlideDuration, long VertResolution, long FramesPerSecond, long Quality)
    {
        static BYTE parms[] = VTS_BSTR VTS_BOOL VTS_I4 VTS_I4 VTS_I4 VTS_I4;
        InvokeHelper(0x84b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality);
    }
    void ApplyTemplate2(LPCTSTR FileName, LPCTSTR VariantGUID)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR;
        InvokeHelper(0x84c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, VariantGUID);
    }
    BOOL get_ChartDataPointTrack()
    {
        BOOL result;
        InvokeHelper(0x84d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ChartDataPointTrack(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x84d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void ExportAsFixedFormat2(LPCTSTR Path, long FixedFormatType, long Intent, long FrameSlides, long HandoutOrder, long OutputType, long PrintHiddenSlides, LPDISPATCH PrintRange, long RangeType, LPCTSTR SlideShowName, BOOL IncludeDocProperties, BOOL KeepIRMSettings, BOOL DocStructureTags, BOOL BitmapMissingFonts, BOOL UseISO19005_1, BOOL IncludeMarkup, VARIANT& ExternalExporter)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_DISPATCH VTS_I4 VTS_BSTR VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_VARIANT;
        InvokeHelper(0x84e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Path, FixedFormatType, Intent, FrameSlides, HandoutOrder, OutputType, PrintHiddenSlides, PrintRange, RangeType, SlideShowName, IncludeDocProperties, KeepIRMSettings, DocStructureTags, BitmapMissingFonts, UseISO19005_1, IncludeMarkup, &ExternalExporter);
    }
    LPDISPATCH get_Guides()
    {
        LPDISPATCH result;
        InvokeHelper(0x84f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }

    // _Presentation 属性
public:

};

// CSlides 包装器类
class CSlides : public COleDispatchDriver
{
public:
    CSlides(){} // 调用 COleDispatchDriver 默认构造函数
    CSlides(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
    CSlides(const CSlides& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

    // 特性
public:

    // 操作
public:


    // Slides 方法
public:
    LPUNKNOWN get__NewEnum()
    {
        LPUNKNOWN result;
        InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
        return result;
    }
    VARIANT _Index(long Index)
    {
        VARIANT result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0xa, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, Index);
        return result;
    }
    long get_Count()
    {
        long result;
        InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Application()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Parent()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH Item(VARIANT& Index)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_VARIANT;
        InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
        return result;
    }
    LPDISPATCH FindBySlideID(long SlideID)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7d3, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, SlideID);
        return result;
    }
    LPDISPATCH Add(long Index, long Layout)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4 VTS_I4;
        InvokeHelper(0x7d4, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index, Layout);
        return result;
    }
    long InsertFromFile(LPCTSTR FileName, long Index, long SlideStart, long SlideEnd)
    {
        long result;
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4 VTS_I4;
        InvokeHelper(0x7d5, DISPATCH_METHOD, VT_I4, (void*)&result, parms, FileName, Index, SlideStart, SlideEnd);
        return result;
    }
    LPDISPATCH Range(VARIANT& Index)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_VARIANT;
        InvokeHelper(0x7d6, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
        return result;
    }
    LPDISPATCH Paste(long Index)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7d7, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
        return result;
    }
    LPDISPATCH AddSlide(long Index, LPDISPATCH pCustomLayout)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4 VTS_DISPATCH;
        InvokeHelper(0x7d8, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index, pCustomLayout);
        return result;
    }

    // Slides 属性
public:

};
// CSlide 包装器类
class CSlide : public COleDispatchDriver
{
public:
    CSlide(){} // 调用 COleDispatchDriver 默认构造函数
    CSlide(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
    CSlide(const CSlide& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

    // 特性
public:

    // 操作
public:


    // _Slide 方法
public:
    LPDISPATCH get_Application()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Parent()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Shapes()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_HeadersFooters()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SlideShowTransition()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ColorScheme()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_ColorScheme(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x7d6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Background()
    {
        LPDISPATCH result;
        InvokeHelper(0x7d7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_Name()
    {
        CString result;
        InvokeHelper(0x7d8, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_Name(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7d8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_SlideID()
    {
        long result;
        InvokeHelper(0x7d9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_PrintSteps()
    {
        long result;
        InvokeHelper(0x7da, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void Select()
    {
        InvokeHelper(0x7db, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Cut()
    {
        InvokeHelper(0x7dc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Copy()
    {
        InvokeHelper(0x7dd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    long get_Layout()
    {
        long result;
        InvokeHelper(0x7de, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Layout(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7de, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH Duplicate()
    {
        LPDISPATCH result;
        InvokeHelper(0x7df, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void Delete()
    {
        InvokeHelper(0x7e0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_Tags()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_SlideIndex()
    {
        long result;
        InvokeHelper(0x7e2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_SlideNumber()
    {
        long result;
        InvokeHelper(0x7e3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_DisplayMasterShapes()
    {
        long result;
        InvokeHelper(0x7e4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_DisplayMasterShapes(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7e4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_FollowMasterBackground()
    {
        long result;
        InvokeHelper(0x7e5, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_FollowMasterBackground(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7e5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_NotesPage()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Master()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Hyperlinks()
    {
        LPDISPATCH result;
        InvokeHelper(0x7e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void Export(LPCTSTR FileName, LPCTSTR FilterName, long ScaleWidth, long ScaleHeight)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_I4 VTS_I4;
        InvokeHelper(0x7e9, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, FilterName, ScaleWidth, ScaleHeight);
    }
    LPDISPATCH get_Scripts()
    {
        LPDISPATCH result;
        InvokeHelper(0x7ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Comments()
    {
        LPDISPATCH result;
        InvokeHelper(0x7ec, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Design()
    {
        LPDISPATCH result;
        InvokeHelper(0x7ed, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_Design(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x7ed, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void MoveTo(long toPos)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7ee, DISPATCH_METHOD, VT_EMPTY, NULL, parms, toPos);
    }
    LPDISPATCH get_TimeLine()
    {
        LPDISPATCH result;
        InvokeHelper(0x7ef, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ApplyTemplate(LPCTSTR FileName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7f0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
    }
    long get_SectionNumber()
    {
        long result;
        InvokeHelper(0x7f1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_CustomLayout()
    {
        LPDISPATCH result;
        InvokeHelper(0x7f2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_CustomLayout(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x7f2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void ApplyTheme(LPCTSTR themeName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7f3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, themeName);
    }
    LPDISPATCH get_ThemeColorScheme()
    {
        LPDISPATCH result;
        InvokeHelper(0x7f4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ApplyThemeColorScheme(LPCTSTR themeColorSchemeName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7f5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, themeColorSchemeName);
    }
    long get_BackgroundStyle()
    {
        long result;
        InvokeHelper(0x7f6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_BackgroundStyle(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7f6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_CustomerData()
    {
        LPDISPATCH result;
        InvokeHelper(0x7f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void PublishSlides(LPCTSTR SlideLibraryUrl, BOOL Overwrite, BOOL UseSlideOrder)
    {
        static BYTE parms[] = VTS_BSTR VTS_BOOL VTS_BOOL;
        InvokeHelper(0x7f8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SlideLibraryUrl, Overwrite, UseSlideOrder);
    }
    void MoveToSectionStart(long toSection)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x7f9, DISPATCH_METHOD, VT_EMPTY, NULL, parms, toSection);
    }
    long get_sectionIndex()
    {
        long result;
        InvokeHelper(0x7fa, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_HasNotesPage()
    {
        long result;
        InvokeHelper(0x7fb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void ApplyTemplate2(LPCTSTR FileName, LPCTSTR VariantGUID)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR;
        InvokeHelper(0x7fc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, VariantGUID);
    }

    // _Slide 属性
public:

};

}

#endif // _PPT_INTERFACES_H_// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类
