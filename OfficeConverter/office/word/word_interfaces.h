#ifndef _WORD_INTERFACES_H_
#define _WORD_INTERFACES_H_

// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

namespace word
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
    LPDISPATCH get_Application()
    {
        LPDISPATCH result;
        InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Creator()
    {
        long result;
        InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Parent()
    {
        LPDISPATCH result;
        InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_Name()
    {
        CString result;
        InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Documents()
    {
        LPDISPATCH result;
        InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Windows()
    {
        LPDISPATCH result;
        InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ActiveDocument()
    {
        LPDISPATCH result;
        InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ActiveWindow()
    {
        LPDISPATCH result;
        InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Selection()
    {
        LPDISPATCH result;
        InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_WordBasic()
    {
        LPDISPATCH result;
        InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_RecentFiles()
    {
        LPDISPATCH result;
        InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_NormalTemplate()
    {
        LPDISPATCH result;
        InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_System()
    {
        LPDISPATCH result;
        InvokeHelper(0x9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_AutoCorrect()
    {
        LPDISPATCH result;
        InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FontNames()
    {
        LPDISPATCH result;
        InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_LandscapeFontNames()
    {
        LPDISPATCH result;
        InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_PortraitFontNames()
    {
        LPDISPATCH result;
        InvokeHelper(0xd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Languages()
    {
        LPDISPATCH result;
        InvokeHelper(0xe, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Assistant()
    {
        LPDISPATCH result;
        InvokeHelper(0xf, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Browser()
    {
        LPDISPATCH result;
        InvokeHelper(0x10, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FileConverters()
    {
        LPDISPATCH result;
        InvokeHelper(0x11, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_MailingLabel()
    {
        LPDISPATCH result;
        InvokeHelper(0x12, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Dialogs()
    {
        LPDISPATCH result;
        InvokeHelper(0x13, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_CaptionLabels()
    {
        LPDISPATCH result;
        InvokeHelper(0x14, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_AutoCaptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x15, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_AddIns()
    {
        LPDISPATCH result;
        InvokeHelper(0x16, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_Visible()
    {
        BOOL result;
        InvokeHelper(0x17, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_Visible(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x17, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_Version()
    {
        CString result;
        InvokeHelper(0x18, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    BOOL get_ScreenUpdating()
    {
        BOOL result;
        InvokeHelper(0x1a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ScreenUpdating(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_PrintPreview()
    {
        BOOL result;
        InvokeHelper(0x1b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_PrintPreview(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Tasks()
    {
        LPDISPATCH result;
        InvokeHelper(0x1c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_DisplayStatusBar()
    {
        BOOL result;
        InvokeHelper(0x1d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DisplayStatusBar(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_SpecialMode()
    {
        BOOL result;
        InvokeHelper(0x1e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    long get_UsableWidth()
    {
        long result;
        InvokeHelper(0x21, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_UsableHeight()
    {
        long result;
        InvokeHelper(0x22, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    BOOL get_MathCoprocessorAvailable()
    {
        BOOL result;
        InvokeHelper(0x24, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    BOOL get_MouseAvailable()
    {
        BOOL result;
        InvokeHelper(0x25, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    VARIANT get_International(long Index)
    {
        VARIANT result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x2e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms, Index);
        return result;
    }
    CString get_Build()
    {
        CString result;
        InvokeHelper(0x2f, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    BOOL get_CapsLock()
    {
        BOOL result;
        InvokeHelper(0x30, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    BOOL get_NumLock()
    {
        BOOL result;
        InvokeHelper(0x31, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    CString get_UserName()
    {
        CString result;
        InvokeHelper(0x34, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_UserName(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x34, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_UserInitials()
    {
        CString result;
        InvokeHelper(0x35, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_UserInitials(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x35, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_UserAddress()
    {
        CString result;
        InvokeHelper(0x36, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_UserAddress(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x36, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_MacroContainer()
    {
        LPDISPATCH result;
        InvokeHelper(0x37, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_DisplayRecentFiles()
    {
        BOOL result;
        InvokeHelper(0x38, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DisplayRecentFiles(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x38, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_CommandBars()
    {
        LPDISPATCH result;
        InvokeHelper(0x39, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SynonymInfo(LPCTSTR Word, VARIANT * LanguageID)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT;
        InvokeHelper(0x3b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, Word, LanguageID);
        return result;
    }
    LPDISPATCH get_VBE()
    {
        LPDISPATCH result;
        InvokeHelper(0x3d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_DefaultSaveFormat()
    {
        CString result;
        InvokeHelper(0x40, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_DefaultSaveFormat(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x40, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_ListGalleries()
    {
        LPDISPATCH result;
        InvokeHelper(0x41, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_ActivePrinter()
    {
        CString result;
        InvokeHelper(0x42, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_ActivePrinter(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x42, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Templates()
    {
        LPDISPATCH result;
        InvokeHelper(0x43, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_CustomizationContext()
    {
        LPDISPATCH result;
        InvokeHelper(0x44, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_CustomizationContext(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x44, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_KeyBindings()
    {
        LPDISPATCH result;
        InvokeHelper(0x45, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_KeysBoundTo(long KeyCategory, LPCTSTR Command, VARIANT * CommandParameter)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4 VTS_BSTR VTS_PVARIANT;
        InvokeHelper(0x46, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, KeyCategory, Command, CommandParameter);
        return result;
    }
    LPDISPATCH get_FindKey(long KeyCode, VARIANT * KeyCode2)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4 VTS_PVARIANT;
        InvokeHelper(0x47, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, KeyCode, KeyCode2);
        return result;
    }
    CString get_Caption()
    {
        CString result;
        InvokeHelper(0x50, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_Caption(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x50, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_Path()
    {
        CString result;
        InvokeHelper(0x51, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    BOOL get_DisplayScrollBars()
    {
        BOOL result;
        InvokeHelper(0x52, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DisplayScrollBars(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x52, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_StartupPath()
    {
        CString result;
        InvokeHelper(0x53, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_StartupPath(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x53, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_BackgroundSavingStatus()
    {
        long result;
        InvokeHelper(0x55, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_BackgroundPrintingStatus()
    {
        long result;
        InvokeHelper(0x56, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_Left()
    {
        long result;
        InvokeHelper(0x57, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Left(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x57, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Top()
    {
        long result;
        InvokeHelper(0x58, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Top(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x58, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Width()
    {
        long result;
        InvokeHelper(0x59, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Width(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x59, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Height()
    {
        long result;
        InvokeHelper(0x5a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Height(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x5a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_WindowState()
    {
        long result;
        InvokeHelper(0x5b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_WindowState(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x5b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_DisplayAutoCompleteTips()
    {
        BOOL result;
        InvokeHelper(0x5c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DisplayAutoCompleteTips(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x5c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Options()
    {
        LPDISPATCH result;
        InvokeHelper(0x5d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_DisplayAlerts()
    {
        long result;
        InvokeHelper(0x5e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_DisplayAlerts(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x5e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_CustomDictionaries()
    {
        LPDISPATCH result;
        InvokeHelper(0x5f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_PathSeparator()
    {
        CString result;
        InvokeHelper(0x60, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_StatusBar(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x61, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_MAPIAvailable()
    {
        BOOL result;
        InvokeHelper(0x62, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    BOOL get_DisplayScreenTips()
    {
        BOOL result;
        InvokeHelper(0x63, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DisplayScreenTips(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x63, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_EnableCancelKey()
    {
        long result;
        InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_EnableCancelKey(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x64, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_UserControl()
    {
        BOOL result;
        InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FileSearch()
    {
        LPDISPATCH result;
        InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_MailSystem()
    {
        long result;
        InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    CString get_DefaultTableSeparator()
    {
        CString result;
        InvokeHelper(0x69, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_DefaultTableSeparator(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x69, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_ShowVisualBasicEditor()
    {
        BOOL result;
        InvokeHelper(0x6a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ShowVisualBasicEditor(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x6a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_BrowseExtraFileTypes()
    {
        CString result;
        InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_BrowseExtraFileTypes(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x6c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_IsObjectValid(LPDISPATCH Object)
    {
        BOOL result;
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x6d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, parms, Object);
        return result;
    }
    LPDISPATCH get_HangulHanjaDictionaries()
    {
        LPDISPATCH result;
        InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_MailMessage()
    {
        LPDISPATCH result;
        InvokeHelper(0x15c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_FocusInMailHeader()
    {
        BOOL result;
        InvokeHelper(0x182, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void Quit(VARIANT * SaveChanges, VARIANT * OriginalFormat, VARIANT * RouteDocument)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x451, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SaveChanges, OriginalFormat, RouteDocument);
    }
    void ScreenRefresh()
    {
        InvokeHelper(0x12d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void PrintOutOld(VARIANT * Background, VARIANT * Append, VARIANT * Range, VARIANT * OutputFileName, VARIANT * From, VARIANT * To, VARIANT * Item, VARIANT * Copies, VARIANT * Pages, VARIANT * PageType, VARIANT * PrintToFile, VARIANT * Collate, VARIANT * FileName, VARIANT * ActivePrinterMacGX, VARIANT * ManualDuplexPrint)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x12e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint);
    }
    void LookupNameProperties(LPCTSTR Name)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x12f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name);
    }
    void SubstituteFont(LPCTSTR UnavailableFont, LPCTSTR SubstituteFont)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR;
        InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, parms, UnavailableFont, SubstituteFont);
    }
    BOOL Repeat(VARIANT * Times)
    {
        BOOL result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x131, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Times);
        return result;
    }
    void DDEExecute(long Channel, LPCTSTR Command)
    {
        static BYTE parms[] = VTS_I4 VTS_BSTR;
        InvokeHelper(0x136, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel, Command);
    }
    long DDEInitiate(LPCTSTR App, LPCTSTR Topic)
    {
        long result;
        static BYTE parms[] = VTS_BSTR VTS_BSTR;
        InvokeHelper(0x137, DISPATCH_METHOD, VT_I4, (void*)&result, parms, App, Topic);
        return result;
    }
    void DDEPoke(long Channel, LPCTSTR Item, LPCTSTR Data)
    {
        static BYTE parms[] = VTS_I4 VTS_BSTR VTS_BSTR;
        InvokeHelper(0x138, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel, Item, Data);
    }
    CString DDERequest(long Channel, LPCTSTR Item)
    {
        CString result;
        static BYTE parms[] = VTS_I4 VTS_BSTR;
        InvokeHelper(0x139, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Channel, Item);
        return result;
    }
    void DDETerminate(long Channel)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x13a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel);
    }
    void DDETerminateAll()
    {
        InvokeHelper(0x13b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    long BuildKeyCode(long Arg1, VARIANT * Arg2, VARIANT * Arg3, VARIANT * Arg4)
    {
        long result;
        static BYTE parms[] = VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x13c, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
        return result;
    }
    CString KeyString(long KeyCode, VARIANT * KeyCode2)
    {
        CString result;
        static BYTE parms[] = VTS_I4 VTS_PVARIANT;
        InvokeHelper(0x13d, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, KeyCode, KeyCode2);
        return result;
    }
    void OrganizerCopy(LPCTSTR Source, LPCTSTR Destination, LPCTSTR Name, long Object)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BSTR VTS_I4;
        InvokeHelper(0x13e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Source, Destination, Name, Object);
    }
    void OrganizerDelete(LPCTSTR Source, LPCTSTR Name, long Object)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_I4;
        InvokeHelper(0x13f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Source, Name, Object);
    }
    void OrganizerRename(LPCTSTR Source, LPCTSTR Name, LPCTSTR NewName, long Object)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BSTR VTS_I4;
        InvokeHelper(0x140, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Source, Name, NewName, Object);
    }
    void AddAddress(SAFEARRAY * * TagID, SAFEARRAY * * Value)
    {
        static BYTE parms[] = VTS_UNKNOWN VTS_UNKNOWN;
        InvokeHelper(0x141, DISPATCH_METHOD, VT_EMPTY, NULL, parms, TagID, Value);
    }
    CString GetAddress(VARIANT * Name, VARIANT * AddressProperties, VARIANT * UseAutoText, VARIANT * DisplaySelectDialog, VARIANT * SelectDialog, VARIANT * CheckNamesDialog, VARIANT * RecentAddressesChoice, VARIANT * UpdateRecentAddresses)
    {
        CString result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x142, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Name, AddressProperties, UseAutoText, DisplaySelectDialog, SelectDialog, CheckNamesDialog, RecentAddressesChoice, UpdateRecentAddresses);
        return result;
    }
    BOOL CheckGrammar(LPCTSTR String)
    {
        BOOL result;
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x143, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, String);
        return result;
    }
    BOOL CheckSpelling(LPCTSTR Word, VARIANT * CustomDictionary, VARIANT * IgnoreUppercase, VARIANT * MainDictionary, VARIANT * CustomDictionary2, VARIANT * CustomDictionary3, VARIANT * CustomDictionary4, VARIANT * CustomDictionary5, VARIANT * CustomDictionary6, VARIANT * CustomDictionary7, VARIANT * CustomDictionary8, VARIANT * CustomDictionary9, VARIANT * CustomDictionary10)
    {
        BOOL result;
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x144, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Word, CustomDictionary, IgnoreUppercase, MainDictionary, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10);
        return result;
    }
    void ResetIgnoreAll()
    {
        InvokeHelper(0x146, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH GetSpellingSuggestions(LPCTSTR Word, VARIANT * CustomDictionary, VARIANT * IgnoreUppercase, VARIANT * MainDictionary, VARIANT * SuggestionMode, VARIANT * CustomDictionary2, VARIANT * CustomDictionary3, VARIANT * CustomDictionary4, VARIANT * CustomDictionary5, VARIANT * CustomDictionary6, VARIANT * CustomDictionary7, VARIANT * CustomDictionary8, VARIANT * CustomDictionary9, VARIANT * CustomDictionary10)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x147, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Word, CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10);
        return result;
    }
    void GoBack()
    {
        InvokeHelper(0x148, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Help(VARIANT * HelpType)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x149, DISPATCH_METHOD, VT_EMPTY, NULL, parms, HelpType);
    }
    void AutomaticChange()
    {
        InvokeHelper(0x14a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ShowMe()
    {
        InvokeHelper(0x14b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void HelpTool()
    {
        InvokeHelper(0x14c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH NewWindow()
    {
        LPDISPATCH result;
        InvokeHelper(0x159, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ListCommands(BOOL ListAllCommands)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x15a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ListAllCommands);
    }
    void ShowClipboard()
    {
        InvokeHelper(0x15d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void OnTime(VARIANT * When, LPCTSTR Name, VARIANT * Tolerance)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_BSTR VTS_PVARIANT;
        InvokeHelper(0x15e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, When, Name, Tolerance);
    }
    void NextLetter()
    {
        InvokeHelper(0x15f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    short MountVolume(LPCTSTR Zone, LPCTSTR Server, LPCTSTR Volume, VARIANT * User, VARIANT * UserPassword, VARIANT * VolumePassword)
    {
        short result;
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x161, DISPATCH_METHOD, VT_I2, (void*)&result, parms, Zone, Server, Volume, User, UserPassword, VolumePassword);
        return result;
    }
    CString CleanString(LPCTSTR String)
    {
        CString result;
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x162, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, String);
        return result;
    }
    void SendFax()
    {
        InvokeHelper(0x164, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ChangeFileOpenDirectory(LPCTSTR Path)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x165, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Path);
    }
    void RunOld(LPCTSTR MacroName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x166, DISPATCH_METHOD, VT_EMPTY, NULL, parms, MacroName);
    }
    void GoForward()
    {
        InvokeHelper(0x167, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Move(long Left, long Top)
    {
        static BYTE parms[] = VTS_I4 VTS_I4;
        InvokeHelper(0x168, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Left, Top);
    }
    void Resize(long Width, long Height)
    {
        static BYTE parms[] = VTS_I4 VTS_I4;
        InvokeHelper(0x169, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Width, Height);
    }
    float InchesToPoints(float Inches)
    {
        float result;
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x172, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Inches);
        return result;
    }
    float CentimetersToPoints(float Centimeters)
    {
        float result;
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x173, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Centimeters);
        return result;
    }
    float MillimetersToPoints(float Millimeters)
    {
        float result;
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x174, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Millimeters);
        return result;
    }
    float PicasToPoints(float Picas)
    {
        float result;
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x175, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Picas);
        return result;
    }
    float LinesToPoints(float Lines)
    {
        float result;
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x176, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Lines);
        return result;
    }
    float PointsToInches(float Points)
    {
        float result;
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x17c, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
        return result;
    }
    float PointsToCentimeters(float Points)
    {
        float result;
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x17d, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
        return result;
    }
    float PointsToMillimeters(float Points)
    {
        float result;
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x17e, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
        return result;
    }
    float PointsToPicas(float Points)
    {
        float result;
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x17f, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
        return result;
    }
    float PointsToLines(float Points)
    {
        float result;
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x180, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
        return result;
    }
    void Activate()
    {
        InvokeHelper(0x181, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    float PointsToPixels(float Points, VARIANT * fVertical)
    {
        float result;
        static BYTE parms[] = VTS_R4 VTS_PVARIANT;
        InvokeHelper(0x183, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points, fVertical);
        return result;
    }
    float PixelsToPoints(float Pixels, VARIANT * fVertical)
    {
        float result;
        static BYTE parms[] = VTS_R4 VTS_PVARIANT;
        InvokeHelper(0x184, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Pixels, fVertical);
        return result;
    }
    void KeyboardLatin()
    {
        InvokeHelper(0x190, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void KeyboardBidi()
    {
        InvokeHelper(0x191, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ToggleKeyboard()
    {
        InvokeHelper(0x192, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    long Keyboard(long LangId)
    {
        long result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1be, DISPATCH_METHOD, VT_I4, (void*)&result, parms, LangId);
        return result;
    }
    CString ProductCode()
    {
        CString result;
        InvokeHelper(0x194, DISPATCH_METHOD, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH DefaultWebOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x195, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void DiscussionSupport(VARIANT * Range, VARIANT * cid, VARIANT * piCSE)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x197, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Range, cid, piCSE);
    }
    void SetDefaultTheme(LPCTSTR Name, long DocumentType)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4;
        InvokeHelper(0x19e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, DocumentType);
    }
    CString GetDefaultTheme(long DocumentType)
    {
        CString result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1a0, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, DocumentType);
        return result;
    }
    LPDISPATCH get_EmailOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x185, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Language()
    {
        long result;
        InvokeHelper(0x187, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_COMAddIns()
    {
        LPDISPATCH result;
        InvokeHelper(0x6f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_CheckLanguage()
    {
        BOOL result;
        InvokeHelper(0x70, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_CheckLanguage(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x70, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_LanguageSettings()
    {
        LPDISPATCH result;
        InvokeHelper(0x193, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_Dummy1()
    {
        BOOL result;
        InvokeHelper(0x196, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_AnswerWizard()
    {
        LPDISPATCH result;
        InvokeHelper(0x199, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_FeatureInstall()
    {
        long result;
        InvokeHelper(0x1bf, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_FeatureInstall(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1bf, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void PrintOut2000(VARIANT * Background, VARIANT * Append, VARIANT * Range, VARIANT * OutputFileName, VARIANT * From, VARIANT * To, VARIANT * Item, VARIANT * Copies, VARIANT * Pages, VARIANT * PageType, VARIANT * PrintToFile, VARIANT * Collate, VARIANT * FileName, VARIANT * ActivePrinterMacGX, VARIANT * ManualDuplexPrint, VARIANT * PrintZoomColumn, VARIANT * PrintZoomRow, VARIANT * PrintZoomPaperWidth, VARIANT * PrintZoomPaperHeight)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1bc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight);
    }
    VARIANT Run(LPCTSTR MacroName, VARIANT * varg1, VARIANT * varg2, VARIANT * varg3, VARIANT * varg4, VARIANT * varg5, VARIANT * varg6, VARIANT * varg7, VARIANT * varg8, VARIANT * varg9, VARIANT * varg10, VARIANT * varg11, VARIANT * varg12, VARIANT * varg13, VARIANT * varg14, VARIANT * varg15, VARIANT * varg16, VARIANT * varg17, VARIANT * varg18, VARIANT * varg19, VARIANT * varg20, VARIANT * varg21, VARIANT * varg22, VARIANT * varg23, VARIANT * varg24, VARIANT * varg25, VARIANT * varg26, VARIANT * varg27, VARIANT * varg28, VARIANT * varg29, VARIANT * varg30)
    {
        VARIANT result;
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1bd, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, MacroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29, varg30);
        return result;
    }
    void PrintOut(VARIANT * Background, VARIANT * Append, VARIANT * Range, VARIANT * OutputFileName, VARIANT * From, VARIANT * To, VARIANT * Item, VARIANT * Copies, VARIANT * Pages, VARIANT * PageType, VARIANT * PrintToFile, VARIANT * Collate, VARIANT * FileName, VARIANT * ActivePrinterMacGX, VARIANT * ManualDuplexPrint, VARIANT * PrintZoomColumn, VARIANT * PrintZoomRow, VARIANT * PrintZoomPaperWidth, VARIANT * PrintZoomPaperHeight)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1c0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight);
    }
    long get_AutomationSecurity()
    {
        long result;
        InvokeHelper(0x1c1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_AutomationSecurity(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1c1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_FileDialog(long FileDialogType)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1c2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, FileDialogType);
        return result;
    }
    CString get_EmailTemplate()
    {
        CString result;
        InvokeHelper(0x1c3, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_EmailTemplate(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x1c3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_ShowWindowsInTaskbar()
    {
        BOOL result;
        InvokeHelper(0x1c4, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ShowWindowsInTaskbar(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1c4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_NewDocument()
    {
        LPDISPATCH result;
        InvokeHelper(0x1c6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_ShowStartupDialog()
    {
        BOOL result;
        InvokeHelper(0x1c7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ShowStartupDialog(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1c7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_AutoCorrectEmail()
    {
        LPDISPATCH result;
        InvokeHelper(0x1c8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_TaskPanes()
    {
        LPDISPATCH result;
        InvokeHelper(0x1c9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_DefaultLegalBlackline()
    {
        BOOL result;
        InvokeHelper(0x1cb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DefaultLegalBlackline(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1cb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL Dummy2()
    {
        BOOL result;
        InvokeHelper(0x1ca, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SmartTagRecognizers()
    {
        LPDISPATCH result;
        InvokeHelper(0x1cc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SmartTagTypes()
    {
        LPDISPATCH result;
        InvokeHelper(0x1cd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_XMLNamespaces()
    {
        LPDISPATCH result;
        InvokeHelper(0x1cf, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void PutFocusInMailHeader()
    {
        InvokeHelper(0x1d0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    BOOL get_ArbitraryXMLSupportAvailable()
    {
        BOOL result;
        InvokeHelper(0x1d1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    CString get_BuildFull()
    {
        CString result;
        InvokeHelper(0x1d2, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_BuildFeatureCrew()
    {
        CString result;
        InvokeHelper(0x1d3, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void LoadMasterList(LPCTSTR FileName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x1d5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
    }
    LPDISPATCH CompareDocuments(LPDISPATCH OriginalDocument, LPDISPATCH RevisedDocument, long Destination, long Granularity, BOOL CompareFormatting, BOOL CompareCaseChanges, BOOL CompareWhitespace, BOOL CompareTables, BOOL CompareHeaders, BOOL CompareFootnotes, BOOL CompareTextboxes, BOOL CompareFields, BOOL CompareComments, BOOL CompareMoves, LPCTSTR RevisedAuthor, BOOL IgnoreAllComparisonWarnings)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_I4 VTS_I4 VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BSTR VTS_BOOL;
        InvokeHelper(0x1d6, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, OriginalDocument, RevisedDocument, Destination, Granularity, CompareFormatting, CompareCaseChanges, CompareWhitespace, CompareTables, CompareHeaders, CompareFootnotes, CompareTextboxes, CompareFields, CompareComments, CompareMoves, RevisedAuthor, IgnoreAllComparisonWarnings);
        return result;
    }
    LPDISPATCH MergeDocuments(LPDISPATCH OriginalDocument, LPDISPATCH RevisedDocument, long Destination, long Granularity, BOOL CompareFormatting, BOOL CompareCaseChanges, BOOL CompareWhitespace, BOOL CompareTables, BOOL CompareHeaders, BOOL CompareFootnotes, BOOL CompareTextboxes, BOOL CompareFields, BOOL CompareComments, BOOL CompareMoves, LPCTSTR OriginalAuthor, LPCTSTR RevisedAuthor, long FormatFrom)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_I4 VTS_I4 VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BOOL VTS_BSTR VTS_BSTR VTS_I4;
        InvokeHelper(0x1d7, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, OriginalDocument, RevisedDocument, Destination, Granularity, CompareFormatting, CompareCaseChanges, CompareWhitespace, CompareTables, CompareHeaders, CompareFootnotes, CompareTextboxes, CompareFields, CompareComments, CompareMoves, OriginalAuthor, RevisedAuthor, FormatFrom);
        return result;
    }
    LPDISPATCH get_Bibliography()
    {
        LPDISPATCH result;
        InvokeHelper(0x1d8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_ShowStylePreviews()
    {
        BOOL result;
        InvokeHelper(0x1d9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ShowStylePreviews(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1d9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_RestrictLinkedStyles()
    {
        BOOL result;
        InvokeHelper(0x1da, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_RestrictLinkedStyles(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1da, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_OMathAutoCorrect()
    {
        LPDISPATCH result;
        InvokeHelper(0x1db, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_DisplayDocumentInformationPanel()
    {
        BOOL result;
        InvokeHelper(0x1dc, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DisplayDocumentInformationPanel(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1dc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Assistance()
    {
        LPDISPATCH result;
        InvokeHelper(0x1dd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_OpenAttachmentsInFullScreen()
    {
        BOOL result;
        InvokeHelper(0x1de, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_OpenAttachmentsInFullScreen(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1de, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_ActiveEncryptionSession()
    {
        long result;
        InvokeHelper(0x1df, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    BOOL get_DontResetInsertionPointProperties()
    {
        BOOL result;
        InvokeHelper(0x1e0, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DontResetInsertionPointProperties(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1e0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_SmartArtLayouts()
    {
        LPDISPATCH result;
        InvokeHelper(0x1e1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SmartArtQuickStyles()
    {
        LPDISPATCH result;
        InvokeHelper(0x1e2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SmartArtColors()
    {
        LPDISPATCH result;
        InvokeHelper(0x1e3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ThreeWayMerge(LPDISPATCH LocalDocument, LPDISPATCH ServerDocument, LPDISPATCH BaseDocument, BOOL FavorSource)
    {
        static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_DISPATCH VTS_BOOL;
        InvokeHelper(0x1e4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, LocalDocument, ServerDocument, BaseDocument, FavorSource);
    }
    void Dummy4()
    {
        InvokeHelper(0x1e5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_UndoRecord()
    {
        LPDISPATCH result;
        InvokeHelper(0x1e6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_PickerDialog()
    {
        LPDISPATCH result;
        InvokeHelper(0x1e9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ProtectedViewWindows()
    {
        LPDISPATCH result;
        InvokeHelper(0x1ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ActiveProtectedViewWindow()
    {
        LPDISPATCH result;
        InvokeHelper(0x1eb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_IsSandboxed()
    {
        BOOL result;
        InvokeHelper(0x1ec, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    long get_FileValidation()
    {
        long result;
        InvokeHelper(0x1ed, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_FileValidation(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1ed, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_ChartDataPointTrack()
    {
        BOOL result;
        InvokeHelper(0x1ee, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ChartDataPointTrack(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1ee, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_ShowAnimation()
    {
        BOOL result;
        InvokeHelper(0x1ef, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ShowAnimation(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1ef, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }

    // _Application 属性
public:

};


// CDocuments 包装器类

class CDocuments : public COleDispatchDriver
{
public:
    CDocuments(){} // 调用 COleDispatchDriver 默认构造函数
    CDocuments(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
    CDocuments(const CDocuments& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

    // 特性
public:

    // 操作
public:


    // Documents 方法
public:
    LPUNKNOWN get__NewEnum()
    {
        LPUNKNOWN result;
        InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
        return result;
    }
    long get_Count()
    {
        long result;
        InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Application()
    {
        LPDISPATCH result;
        InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Creator()
    {
        long result;
        InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Parent()
    {
        LPDISPATCH result;
        InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH Item(VARIANT * Index)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
        return result;
    }
    void Close(VARIANT * SaveChanges, VARIANT * OriginalFormat, VARIANT * RouteDocument)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x451, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SaveChanges, OriginalFormat, RouteDocument);
    }
    LPDISPATCH AddOld(VARIANT * Template, VARIANT * NewTemplate)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xb, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Template, NewTemplate);
        return result;
    }
    LPDISPATCH OpenOld(VARIANT * FileName, VARIANT * ConfirmConversions, VARIANT * ReadOnly, VARIANT * AddToRecentFiles, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * Revert, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate, VARIANT * Format)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xc, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format);
        return result;
    }
    void Save(VARIANT * NoPrompt, VARIANT * OriginalFormat)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NoPrompt, OriginalFormat);
    }
    LPDISPATCH Add(VARIANT * Template, VARIANT * NewTemplate, VARIANT * DocumentType, VARIANT * Visible)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xe, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Template, NewTemplate, DocumentType, Visible);
        return result;
    }
    LPDISPATCH Open2000(VARIANT * FileName, VARIANT * ConfirmConversions, VARIANT * ReadOnly, VARIANT * AddToRecentFiles, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * Revert, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate, VARIANT * Format, VARIANT * Encoding, VARIANT * Visible)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xf, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible);
        return result;
    }
    void CheckOut(LPCTSTR FileName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x10, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
    }
    BOOL CanCheckOut(LPCTSTR FileName)
    {
        BOOL result;
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x11, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, FileName);
        return result;
    }
    LPDISPATCH Open2002(VARIANT * FileName, VARIANT * ConfirmConversions, VARIANT * ReadOnly, VARIANT * AddToRecentFiles, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * Revert, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate, VARIANT * Format, VARIANT * Encoding, VARIANT * Visible, VARIANT * OpenAndRepair, VARIANT * DocumentDirection, VARIANT * NoEncodingDialog)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x12, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog);
        return result;
    }
    LPDISPATCH Open(VARIANT * FileName, VARIANT * ConfirmConversions, VARIANT * ReadOnly, VARIANT * AddToRecentFiles, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * Revert, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate, VARIANT * Format, VARIANT * Encoding, VARIANT * Visible, VARIANT * OpenAndRepair, VARIANT * DocumentDirection, VARIANT * NoEncodingDialog, VARIANT * XMLTransform)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x13, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform);
        return result;
    }
    LPDISPATCH OpenNoRepairDialog(VARIANT * FileName, VARIANT * ConfirmConversions, VARIANT * ReadOnly, VARIANT * AddToRecentFiles, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * Revert, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate, VARIANT * Format, VARIANT * Encoding, VARIANT * Visible, VARIANT * OpenAndRepair, VARIANT * DocumentDirection, VARIANT * NoEncodingDialog, VARIANT * XMLTransform)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x14, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform);
        return result;
    }
    LPDISPATCH AddBlogDocument(LPCTSTR ProviderID, LPCTSTR PostURL, LPCTSTR BlogName, LPCTSTR PostID)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR;
        InvokeHelper(0x15, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, ProviderID, PostURL, BlogName, PostID);
        return result;
    }

    // Documents 属性
public:

};



// CDocument 包装器类

class CDocument : public COleDispatchDriver
{
public:
    CDocument(){} // 调用 COleDispatchDriver 默认构造函数
    CDocument(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
    CDocument(const CDocument& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

    // 特性
public:

    // 操作
public:


    // _Document 方法
public:
    CString get_Name()
    {
        CString result;
        InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Application()
    {
        LPDISPATCH result;
        InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Creator()
    {
        long result;
        InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Parent()
    {
        LPDISPATCH result;
        InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_BuiltInDocumentProperties()
    {
        LPDISPATCH result;
        InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_CustomDocumentProperties()
    {
        LPDISPATCH result;
        InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_Path()
    {
        CString result;
        InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Bookmarks()
    {
        LPDISPATCH result;
        InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Tables()
    {
        LPDISPATCH result;
        InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Footnotes()
    {
        LPDISPATCH result;
        InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Endnotes()
    {
        LPDISPATCH result;
        InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Comments()
    {
        LPDISPATCH result;
        InvokeHelper(0x9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Type()
    {
        long result;
        InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    BOOL get_AutoHyphenation()
    {
        BOOL result;
        InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_AutoHyphenation(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0xb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_HyphenateCaps()
    {
        BOOL result;
        InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_HyphenateCaps(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0xc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_HyphenationZone()
    {
        long result;
        InvokeHelper(0xd, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_HyphenationZone(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0xd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_ConsecutiveHyphensLimit()
    {
        long result;
        InvokeHelper(0xe, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_ConsecutiveHyphensLimit(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0xe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Sections()
    {
        LPDISPATCH result;
        InvokeHelper(0xf, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Paragraphs()
    {
        LPDISPATCH result;
        InvokeHelper(0x10, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Words()
    {
        LPDISPATCH result;
        InvokeHelper(0x11, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Sentences()
    {
        LPDISPATCH result;
        InvokeHelper(0x12, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Characters()
    {
        LPDISPATCH result;
        InvokeHelper(0x13, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Fields()
    {
        LPDISPATCH result;
        InvokeHelper(0x14, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FormFields()
    {
        LPDISPATCH result;
        InvokeHelper(0x15, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Styles()
    {
        LPDISPATCH result;
        InvokeHelper(0x16, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Frames()
    {
        LPDISPATCH result;
        InvokeHelper(0x17, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_TablesOfFigures()
    {
        LPDISPATCH result;
        InvokeHelper(0x19, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Variables()
    {
        LPDISPATCH result;
        InvokeHelper(0x1a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_MailMerge()
    {
        LPDISPATCH result;
        InvokeHelper(0x1b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Envelope()
    {
        LPDISPATCH result;
        InvokeHelper(0x1c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_FullName()
    {
        CString result;
        InvokeHelper(0x1d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Revisions()
    {
        LPDISPATCH result;
        InvokeHelper(0x1e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_TablesOfContents()
    {
        LPDISPATCH result;
        InvokeHelper(0x1f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_TablesOfAuthorities()
    {
        LPDISPATCH result;
        InvokeHelper(0x20, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_PageSetup()
    {
        LPDISPATCH result;
        InvokeHelper(0x44d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_PageSetup(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x44d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Windows()
    {
        LPDISPATCH result;
        InvokeHelper(0x22, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_HasRoutingSlip()
    {
        BOOL result;
        InvokeHelper(0x23, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_HasRoutingSlip(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x23, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_RoutingSlip()
    {
        LPDISPATCH result;
        InvokeHelper(0x24, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_Routed()
    {
        BOOL result;
        InvokeHelper(0x25, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_TablesOfAuthoritiesCategories()
    {
        LPDISPATCH result;
        InvokeHelper(0x26, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Indexes()
    {
        LPDISPATCH result;
        InvokeHelper(0x27, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_Saved()
    {
        BOOL result;
        InvokeHelper(0x28, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_Saved(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x28, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Content()
    {
        LPDISPATCH result;
        InvokeHelper(0x29, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ActiveWindow()
    {
        LPDISPATCH result;
        InvokeHelper(0x2a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Kind()
    {
        long result;
        InvokeHelper(0x2b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Kind(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x2b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_ReadOnly()
    {
        BOOL result;
        InvokeHelper(0x2c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Subdocuments()
    {
        LPDISPATCH result;
        InvokeHelper(0x2d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_IsMasterDocument()
    {
        BOOL result;
        InvokeHelper(0x2e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    float get_DefaultTabStop()
    {
        float result;
        InvokeHelper(0x30, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_DefaultTabStop(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x30, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_EmbedTrueTypeFonts()
    {
        BOOL result;
        InvokeHelper(0x32, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_EmbedTrueTypeFonts(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x32, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_SaveFormsData()
    {
        BOOL result;
        InvokeHelper(0x33, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_SaveFormsData(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x33, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_ReadOnlyRecommended()
    {
        BOOL result;
        InvokeHelper(0x34, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ReadOnlyRecommended(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x34, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_SaveSubsetFonts()
    {
        BOOL result;
        InvokeHelper(0x35, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_SaveSubsetFonts(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x35, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_Compatibility(long Type)
    {
        BOOL result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x37, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, parms, Type);
        return result;
    }
    void put_Compatibility(long Type, BOOL newValue)
    {
        static BYTE parms[] = VTS_I4 VTS_BOOL;
        InvokeHelper(0x37, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, Type, newValue);
    }
    LPDISPATCH get_StoryRanges()
    {
        LPDISPATCH result;
        InvokeHelper(0x38, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_CommandBars()
    {
        LPDISPATCH result;
        InvokeHelper(0x39, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_IsSubdocument()
    {
        BOOL result;
        InvokeHelper(0x3a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    long get_SaveFormat()
    {
        long result;
        InvokeHelper(0x3b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_ProtectionType()
    {
        long result;
        InvokeHelper(0x3c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Hyperlinks()
    {
        LPDISPATCH result;
        InvokeHelper(0x3d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Shapes()
    {
        LPDISPATCH result;
        InvokeHelper(0x3e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ListTemplates()
    {
        LPDISPATCH result;
        InvokeHelper(0x3f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Lists()
    {
        LPDISPATCH result;
        InvokeHelper(0x40, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_UpdateStylesOnOpen()
    {
        BOOL result;
        InvokeHelper(0x42, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_UpdateStylesOnOpen(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x42, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    VARIANT get_AttachedTemplate()
    {
        VARIANT result;
        InvokeHelper(0x43, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    void put_AttachedTemplate(VARIANT * newValue)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x43, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_InlineShapes()
    {
        LPDISPATCH result;
        InvokeHelper(0x44, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Background()
    {
        LPDISPATCH result;
        InvokeHelper(0x45, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_Background(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x45, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_GrammarChecked()
    {
        BOOL result;
        InvokeHelper(0x46, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_GrammarChecked(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x46, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_SpellingChecked()
    {
        BOOL result;
        InvokeHelper(0x47, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_SpellingChecked(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x47, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_ShowGrammaticalErrors()
    {
        BOOL result;
        InvokeHelper(0x48, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ShowGrammaticalErrors(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x48, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_ShowSpellingErrors()
    {
        BOOL result;
        InvokeHelper(0x49, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ShowSpellingErrors(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x49, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Versions()
    {
        LPDISPATCH result;
        InvokeHelper(0x4b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_ShowSummary()
    {
        BOOL result;
        InvokeHelper(0x4c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ShowSummary(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x4c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_SummaryViewMode()
    {
        long result;
        InvokeHelper(0x4d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_SummaryViewMode(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x4d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_SummaryLength()
    {
        long result;
        InvokeHelper(0x4e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_SummaryLength(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x4e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_PrintFractionalWidths()
    {
        BOOL result;
        InvokeHelper(0x4f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_PrintFractionalWidths(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x4f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_PrintPostScriptOverText()
    {
        BOOL result;
        InvokeHelper(0x50, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_PrintPostScriptOverText(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x50, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Container()
    {
        LPDISPATCH result;
        InvokeHelper(0x52, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_PrintFormsData()
    {
        BOOL result;
        InvokeHelper(0x53, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_PrintFormsData(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x53, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_ListParagraphs()
    {
        LPDISPATCH result;
        InvokeHelper(0x54, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_Password(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x55, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void put_WritePassword(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x56, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_HasPassword()
    {
        BOOL result;
        InvokeHelper(0x57, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    BOOL get_WriteReserved()
    {
        BOOL result;
        InvokeHelper(0x58, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    CString get_ActiveWritingStyle(VARIANT * LanguageID)
    {
        CString result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x5a, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, LanguageID);
        return result;
    }
    void put_ActiveWritingStyle(VARIANT * LanguageID, LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_BSTR;
        InvokeHelper(0x5a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, LanguageID, newValue);
    }
    BOOL get_UserControl()
    {
        BOOL result;
        InvokeHelper(0x5c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_UserControl(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x5c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_HasMailer()
    {
        BOOL result;
        InvokeHelper(0x5d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_HasMailer(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x5d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Mailer()
    {
        LPDISPATCH result;
        InvokeHelper(0x5e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ReadabilityStatistics()
    {
        LPDISPATCH result;
        InvokeHelper(0x60, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_GrammaticalErrors()
    {
        LPDISPATCH result;
        InvokeHelper(0x61, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SpellingErrors()
    {
        LPDISPATCH result;
        InvokeHelper(0x62, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_VBProject()
    {
        LPDISPATCH result;
        InvokeHelper(0x63, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_FormsDesign()
    {
        BOOL result;
        InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    CString get__CodeName()
    {
        CString result;
        InvokeHelper(0x80010000, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put__CodeName(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x80010000, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_CodeName()
    {
        CString result;
        InvokeHelper(0x106, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    BOOL get_SnapToGrid()
    {
        BOOL result;
        InvokeHelper(0x12c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_SnapToGrid(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x12c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_SnapToShapes()
    {
        BOOL result;
        InvokeHelper(0x12d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_SnapToShapes(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x12d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_GridDistanceHorizontal()
    {
        float result;
        InvokeHelper(0x12e, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_GridDistanceHorizontal(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x12e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_GridDistanceVertical()
    {
        float result;
        InvokeHelper(0x12f, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_GridDistanceVertical(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x12f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_GridOriginHorizontal()
    {
        float result;
        InvokeHelper(0x130, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_GridOriginHorizontal(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x130, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_GridOriginVertical()
    {
        float result;
        InvokeHelper(0x131, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_GridOriginVertical(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x131, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_GridSpaceBetweenHorizontalLines()
    {
        long result;
        InvokeHelper(0x132, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_GridSpaceBetweenHorizontalLines(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x132, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_GridSpaceBetweenVerticalLines()
    {
        long result;
        InvokeHelper(0x133, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_GridSpaceBetweenVerticalLines(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x133, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_GridOriginFromMargin()
    {
        BOOL result;
        InvokeHelper(0x134, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_GridOriginFromMargin(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x134, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_KerningByAlgorithm()
    {
        BOOL result;
        InvokeHelper(0x135, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_KerningByAlgorithm(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x135, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_JustificationMode()
    {
        long result;
        InvokeHelper(0x136, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_JustificationMode(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x136, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_FarEastLineBreakLevel()
    {
        long result;
        InvokeHelper(0x137, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_FarEastLineBreakLevel(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x137, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_NoLineBreakBefore()
    {
        CString result;
        InvokeHelper(0x138, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_NoLineBreakBefore(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x138, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_NoLineBreakAfter()
    {
        CString result;
        InvokeHelper(0x139, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_NoLineBreakAfter(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x139, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_TrackRevisions()
    {
        BOOL result;
        InvokeHelper(0x13a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_TrackRevisions(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x13a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_PrintRevisions()
    {
        BOOL result;
        InvokeHelper(0x13b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_PrintRevisions(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x13b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_ShowRevisions()
    {
        BOOL result;
        InvokeHelper(0x13c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ShowRevisions(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x13c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void Close(VARIANT * SaveChanges, VARIANT * OriginalFormat, VARIANT * RouteDocument)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x451, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SaveChanges, OriginalFormat, RouteDocument);
    }
    void SaveAs2000(VARIANT * FileName, VARIANT * FileFormat, VARIANT * LockComments, VARIANT * Password, VARIANT * AddToRecentFiles, VARIANT * WritePassword, VARIANT * ReadOnlyRecommended, VARIANT * EmbedTrueTypeFonts, VARIANT * SaveNativePictureFormat, VARIANT * SaveFormsData, VARIANT * SaveAsAOCELetter)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter);
    }
    void Repaginate()
    {
        InvokeHelper(0x67, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void FitToPages()
    {
        InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ManualHyphenation()
    {
        InvokeHelper(0x69, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Select()
    {
        InvokeHelper(0xffff, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void DataForm()
    {
        InvokeHelper(0x6a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Route()
    {
        InvokeHelper(0x6b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Save()
    {
        InvokeHelper(0x6c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void PrintOutOld(VARIANT * Background, VARIANT * Append, VARIANT * Range, VARIANT * OutputFileName, VARIANT * From, VARIANT * To, VARIANT * Item, VARIANT * Copies, VARIANT * Pages, VARIANT * PageType, VARIANT * PrintToFile, VARIANT * Collate, VARIANT * ActivePrinterMacGX, VARIANT * ManualDuplexPrint)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint);
    }
    void SendMail()
    {
        InvokeHelper(0x6e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH Range(VARIANT * Start, VARIANT * End)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x7d0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Start, End);
        return result;
    }
    void RunAutoMacro(long Which)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x70, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Which);
    }
    void Activate()
    {
        InvokeHelper(0x71, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void PrintPreview()
    {
        InvokeHelper(0x72, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH GoTo(VARIANT * What, VARIANT * Which, VARIANT * Count, VARIANT * Name)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x73, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, What, Which, Count, Name);
        return result;
    }
    BOOL Undo(VARIANT * Times)
    {
        BOOL result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x74, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Times);
        return result;
    }
    BOOL Redo(VARIANT * Times)
    {
        BOOL result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x75, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Times);
        return result;
    }
    long ComputeStatistics(long Statistic, VARIANT * IncludeFootnotesAndEndnotes)
    {
        long result;
        static BYTE parms[] = VTS_I4 VTS_PVARIANT;
        InvokeHelper(0x76, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Statistic, IncludeFootnotesAndEndnotes);
        return result;
    }
    void MakeCompatibilityDefault()
    {
        InvokeHelper(0x77, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Protect2002(long Type, VARIANT * NoReset, VARIANT * Password)
    {
        static BYTE parms[] = VTS_I4 VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x78, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type, NoReset, Password);
    }
    void Unprotect(VARIANT * Password)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x79, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Password);
    }
    void EditionOptions(long Type, long Option, LPCTSTR Name, VARIANT * Format)
    {
        static BYTE parms[] = VTS_I4 VTS_I4 VTS_BSTR VTS_PVARIANT;
        InvokeHelper(0x7a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type, Option, Name, Format);
    }
    void RunLetterWizard(VARIANT * LetterContent, VARIANT * WizardMode)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x7b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, LetterContent, WizardMode);
    }
    LPDISPATCH GetLetterContent()
    {
        LPDISPATCH result;
        InvokeHelper(0x7c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void SetLetterContent(VARIANT * LetterContent)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x7d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, LetterContent);
    }
    void CopyStylesFromTemplate(LPCTSTR Template)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x7e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Template);
    }
    void UpdateStyles()
    {
        InvokeHelper(0x7f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void CheckGrammar()
    {
        InvokeHelper(0x83, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void CheckSpelling(VARIANT * CustomDictionary, VARIANT * IgnoreUppercase, VARIANT * AlwaysSuggest, VARIANT * CustomDictionary2, VARIANT * CustomDictionary3, VARIANT * CustomDictionary4, VARIANT * CustomDictionary5, VARIANT * CustomDictionary6, VARIANT * CustomDictionary7, VARIANT * CustomDictionary8, VARIANT * CustomDictionary9, VARIANT * CustomDictionary10)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x84, DISPATCH_METHOD, VT_EMPTY, NULL, parms, CustomDictionary, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10);
    }
    void FollowHyperlink(VARIANT * Address, VARIANT * SubAddress, VARIANT * NewWindow, VARIANT * AddHistory, VARIANT * ExtraInfo, VARIANT * Method, VARIANT * HeaderInfo)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x87, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo);
    }
    void AddToFavorites()
    {
        InvokeHelper(0x88, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Reload()
    {
        InvokeHelper(0x89, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH AutoSummarize(VARIANT * Length, VARIANT * Mode, VARIANT * UpdateProperties)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x8a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Length, Mode, UpdateProperties);
        return result;
    }
    void RemoveNumbers(VARIANT * NumberType)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x8c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumberType);
    }
    void ConvertNumbersToText(VARIANT * NumberType)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x8d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumberType);
    }
    long CountNumberedItems(VARIANT * NumberType, VARIANT * Level)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x8e, DISPATCH_METHOD, VT_I4, (void*)&result, parms, NumberType, Level);
        return result;
    }
    void Post()
    {
        InvokeHelper(0x8f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ToggleFormsDesign()
    {
        InvokeHelper(0x90, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Compare2000(LPCTSTR Name)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x91, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name);
    }
    void UpdateSummaryProperties()
    {
        InvokeHelper(0x92, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    VARIANT GetCrossReferenceItems(VARIANT * ReferenceType)
    {
        VARIANT result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x93, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, ReferenceType);
        return result;
    }
    void AutoFormat()
    {
        InvokeHelper(0x94, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ViewCode()
    {
        InvokeHelper(0x95, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ViewPropertyBrowser()
    {
        InvokeHelper(0x96, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ForwardMailer()
    {
        InvokeHelper(0xfa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Reply()
    {
        InvokeHelper(0xfb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ReplyAll()
    {
        InvokeHelper(0xfc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SendMailer(VARIANT * FileFormat, VARIANT * Priority)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xfd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileFormat, Priority);
    }
    void UndoClear()
    {
        InvokeHelper(0xfe, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void PresentIt()
    {
        InvokeHelper(0xff, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SendFax(LPCTSTR Address, VARIANT * Subject)
    {
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT;
        InvokeHelper(0x100, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Address, Subject);
    }
    void Merge2000(LPCTSTR FileName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x101, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
    }
    void ClosePrintPreview()
    {
        InvokeHelper(0x102, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void CheckConsistency()
    {
        InvokeHelper(0x103, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH CreateLetterContent(LPCTSTR DateFormat, BOOL IncludeHeaderFooter, LPCTSTR PageDesign, long LetterStyle, BOOL Letterhead, long LetterheadLocation, float LetterheadSize, LPCTSTR RecipientName, LPCTSTR RecipientAddress, LPCTSTR Salutation, long SalutationType, LPCTSTR RecipientReference, LPCTSTR MailingInstructions, LPCTSTR AttentionLine, LPCTSTR Subject, LPCTSTR CCList, LPCTSTR ReturnAddress, LPCTSTR SenderName, LPCTSTR Closing, LPCTSTR SenderCompany, LPCTSTR SenderJobTitle, LPCTSTR SenderInitials, long EnclosureNumber, VARIANT * InfoBlock, VARIANT * RecipientCode, VARIANT * RecipientGender, VARIANT * ReturnAddressShortForm, VARIANT * SenderCity, VARIANT * SenderCode, VARIANT * SenderGender, VARIANT * SenderReference)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR VTS_BOOL VTS_BSTR VTS_I4 VTS_BOOL VTS_I4 VTS_R4 VTS_BSTR VTS_BSTR VTS_BSTR VTS_I4 VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x104, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, DateFormat, IncludeHeaderFooter, PageDesign, LetterStyle, Letterhead, LetterheadLocation, LetterheadSize, RecipientName, RecipientAddress, Salutation, SalutationType, RecipientReference, MailingInstructions, AttentionLine, Subject, CCList, ReturnAddress, SenderName, Closing, SenderCompany, SenderJobTitle, SenderInitials, EnclosureNumber, InfoBlock, RecipientCode, RecipientGender, ReturnAddressShortForm, SenderCity, SenderCode, SenderGender, SenderReference);
        return result;
    }
    void AcceptAllRevisions()
    {
        InvokeHelper(0x13d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void RejectAllRevisions()
    {
        InvokeHelper(0x13e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void DetectLanguage()
    {
        InvokeHelper(0x97, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ApplyTheme(LPCTSTR Name)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x142, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name);
    }
    void RemoveTheme()
    {
        InvokeHelper(0x143, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void WebPagePreview()
    {
        InvokeHelper(0x145, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ReloadAs(long Encoding)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x14b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Encoding);
    }
    CString get_ActiveTheme()
    {
        CString result;
        InvokeHelper(0x21c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_ActiveThemeDisplayName()
    {
        CString result;
        InvokeHelper(0x21d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Email()
    {
        LPDISPATCH result;
        InvokeHelper(0x13f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Scripts()
    {
        LPDISPATCH result;
        InvokeHelper(0x140, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_LanguageDetected()
    {
        BOOL result;
        InvokeHelper(0x141, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_LanguageDetected(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x141, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_FarEastLineBreakLanguage()
    {
        long result;
        InvokeHelper(0x146, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_FarEastLineBreakLanguage(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x146, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Frameset()
    {
        LPDISPATCH result;
        InvokeHelper(0x147, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    VARIANT get_ClickAndTypeParagraphStyle()
    {
        VARIANT result;
        InvokeHelper(0x148, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    void put_ClickAndTypeParagraphStyle(VARIANT * newValue)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x148, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_HTMLProject()
    {
        LPDISPATCH result;
        InvokeHelper(0x149, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_WebOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x14a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_OpenEncoding()
    {
        long result;
        InvokeHelper(0x14c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_SaveEncoding()
    {
        long result;
        InvokeHelper(0x14d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_SaveEncoding(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x14d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_OptimizeForWord97()
    {
        BOOL result;
        InvokeHelper(0x14e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_OptimizeForWord97(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x14e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_VBASigned()
    {
        BOOL result;
        InvokeHelper(0x14f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void PrintOut2000(VARIANT * Background, VARIANT * Append, VARIANT * Range, VARIANT * OutputFileName, VARIANT * From, VARIANT * To, VARIANT * Item, VARIANT * Copies, VARIANT * Pages, VARIANT * PageType, VARIANT * PrintToFile, VARIANT * Collate, VARIANT * ActivePrinterMacGX, VARIANT * ManualDuplexPrint, VARIANT * PrintZoomColumn, VARIANT * PrintZoomRow, VARIANT * PrintZoomPaperWidth, VARIANT * PrintZoomPaperHeight)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1bc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight);
    }
    void sblt(LPCTSTR s)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x1bd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, s);
    }
    void ConvertVietDoc(long CodePageOrigin)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1bf, DISPATCH_METHOD, VT_EMPTY, NULL, parms, CodePageOrigin);
    }
    void PrintOut(VARIANT * Background, VARIANT * Append, VARIANT * Range, VARIANT * OutputFileName, VARIANT * From, VARIANT * To, VARIANT * Item, VARIANT * Copies, VARIANT * Pages, VARIANT * PageType, VARIANT * PrintToFile, VARIANT * Collate, VARIANT * ActivePrinterMacGX, VARIANT * ManualDuplexPrint, VARIANT * PrintZoomColumn, VARIANT * PrintZoomRow, VARIANT * PrintZoomPaperWidth, VARIANT * PrintZoomPaperHeight)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1be, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight);
    }
    LPDISPATCH get_MailEnvelope()
    {
        LPDISPATCH result;
        InvokeHelper(0x150, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_DisableFeatures()
    {
        BOOL result;
        InvokeHelper(0x151, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DisableFeatures(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x151, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_DoNotEmbedSystemFonts()
    {
        BOOL result;
        InvokeHelper(0x152, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DoNotEmbedSystemFonts(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x152, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Signatures()
    {
        LPDISPATCH result;
        InvokeHelper(0x153, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_DefaultTargetFrame()
    {
        CString result;
        InvokeHelper(0x154, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_DefaultTargetFrame(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x154, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_HTMLDivisions()
    {
        LPDISPATCH result;
        InvokeHelper(0x156, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_DisableFeaturesIntroducedAfter()
    {
        long result;
        InvokeHelper(0x157, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_DisableFeaturesIntroducedAfter(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x157, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_RemovePersonalInformation()
    {
        BOOL result;
        InvokeHelper(0x158, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_RemovePersonalInformation(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x158, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_SmartTags()
    {
        LPDISPATCH result;
        InvokeHelper(0x15a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void Compare2002(LPCTSTR Name, VARIANT * AuthorName, VARIANT * CompareTarget, VARIANT * DetectFormatChanges, VARIANT * IgnoreAllComparisonWarnings, VARIANT * AddToRecentFiles)
    {
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x159, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles);
    }
    void CheckIn(BOOL SaveChanges, VARIANT * Comments, BOOL MakePublic)
    {
        static BYTE parms[] = VTS_BOOL VTS_PVARIANT VTS_BOOL;
        InvokeHelper(0x15d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SaveChanges, Comments, MakePublic);
    }
    BOOL CanCheckin()
    {
        BOOL result;
        InvokeHelper(0x15f, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void Merge(LPCTSTR FileName, VARIANT * MergeTarget, VARIANT * DetectFormatChanges, VARIANT * UseFormattingFrom, VARIANT * AddToRecentFiles)
    {
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x16a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, MergeTarget, DetectFormatChanges, UseFormattingFrom, AddToRecentFiles);
    }
    BOOL get_EmbedSmartTags()
    {
        BOOL result;
        InvokeHelper(0x15b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_EmbedSmartTags(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x15b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_SmartTagsAsXMLProps()
    {
        BOOL result;
        InvokeHelper(0x15c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_SmartTagsAsXMLProps(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x15c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_TextEncoding()
    {
        long result;
        InvokeHelper(0x165, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_TextEncoding(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x165, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_TextLineEnding()
    {
        long result;
        InvokeHelper(0x166, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_TextLineEnding(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x166, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void SendForReview(VARIANT * Recipients, VARIANT * Subject, VARIANT * ShowMessage, VARIANT * IncludeAttachment)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x161, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Recipients, Subject, ShowMessage, IncludeAttachment);
    }
    void ReplyWithChanges(VARIANT * ShowMessage)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x162, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ShowMessage);
    }
    void EndReview()
    {
        InvokeHelper(0x164, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_StyleSheets()
    {
        LPDISPATCH result;
        InvokeHelper(0x168, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    VARIANT get_DefaultTableStyle()
    {
        VARIANT result;
        InvokeHelper(0x16d, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    CString get_PasswordEncryptionProvider()
    {
        CString result;
        InvokeHelper(0x16f, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_PasswordEncryptionAlgorithm()
    {
        CString result;
        InvokeHelper(0x170, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    long get_PasswordEncryptionKeyLength()
    {
        long result;
        InvokeHelper(0x171, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    BOOL get_PasswordEncryptionFileProperties()
    {
        BOOL result;
        InvokeHelper(0x172, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void SetPasswordEncryptionOptions(LPCTSTR PasswordEncryptionProvider, LPCTSTR PasswordEncryptionAlgorithm, long PasswordEncryptionKeyLength, VARIANT * PasswordEncryptionFileProperties)
    {
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_I4 VTS_PVARIANT;
        InvokeHelper(0x169, DISPATCH_METHOD, VT_EMPTY, NULL, parms, PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties);
    }
    void RecheckSmartTags()
    {
        InvokeHelper(0x16b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void RemoveSmartTags()
    {
        InvokeHelper(0x16c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SetDefaultTableStyle(VARIANT * Style, BOOL SetInTemplate)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_BOOL;
        InvokeHelper(0x16e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Style, SetInTemplate);
    }
    void DeleteAllComments()
    {
        InvokeHelper(0x173, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void AcceptAllRevisionsShown()
    {
        InvokeHelper(0x174, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void RejectAllRevisionsShown()
    {
        InvokeHelper(0x175, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void DeleteAllCommentsShown()
    {
        InvokeHelper(0x176, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ResetFormFields()
    {
        InvokeHelper(0x177, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SaveAs(VARIANT * FileName, VARIANT * FileFormat, VARIANT * LockComments, VARIANT * Password, VARIANT * AddToRecentFiles, VARIANT * WritePassword, VARIANT * ReadOnlyRecommended, VARIANT * EmbedTrueTypeFonts, VARIANT * SaveNativePictureFormat, VARIANT * SaveFormsData, VARIANT * SaveAsAOCELetter, VARIANT * Encoding, VARIANT * InsertLineBreaks, VARIANT * AllowSubstitutions, VARIANT * LineEnding, VARIANT * AddBiDiMarks)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x178, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks);
    }
    BOOL get_EmbedLinguisticData()
    {
        BOOL result;
        InvokeHelper(0x179, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_EmbedLinguisticData(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x179, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_FormattingShowFont()
    {
        BOOL result;
        InvokeHelper(0x1c0, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_FormattingShowFont(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1c0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_FormattingShowClear()
    {
        BOOL result;
        InvokeHelper(0x1c1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_FormattingShowClear(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1c1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_FormattingShowParagraph()
    {
        BOOL result;
        InvokeHelper(0x1c2, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_FormattingShowParagraph(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1c2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_FormattingShowNumbering()
    {
        BOOL result;
        InvokeHelper(0x1c3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_FormattingShowNumbering(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1c3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_FormattingShowFilter()
    {
        long result;
        InvokeHelper(0x1c4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_FormattingShowFilter(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1c4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void CheckNewSmartTags()
    {
        InvokeHelper(0x17a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_Permission()
    {
        LPDISPATCH result;
        InvokeHelper(0x1c5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_XMLNodes()
    {
        LPDISPATCH result;
        InvokeHelper(0x1cc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_XMLSchemaReferences()
    {
        LPDISPATCH result;
        InvokeHelper(0x1cd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SmartDocument()
    {
        LPDISPATCH result;
        InvokeHelper(0x1ce, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SharedWorkspace()
    {
        LPDISPATCH result;
        InvokeHelper(0x1cf, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Sync()
    {
        LPDISPATCH result;
        InvokeHelper(0x1d2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_EnforceStyle()
    {
        BOOL result;
        InvokeHelper(0x1d7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_EnforceStyle(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1d7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_AutoFormatOverride()
    {
        BOOL result;
        InvokeHelper(0x1d8, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_AutoFormatOverride(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1d8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_XMLSaveDataOnly()
    {
        BOOL result;
        InvokeHelper(0x1d9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_XMLSaveDataOnly(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1d9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_XMLHideNamespaces()
    {
        BOOL result;
        InvokeHelper(0x1dd, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_XMLHideNamespaces(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1dd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_XMLShowAdvancedErrors()
    {
        BOOL result;
        InvokeHelper(0x1de, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_XMLShowAdvancedErrors(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1de, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_XMLUseXSLTWhenSaving()
    {
        BOOL result;
        InvokeHelper(0x1da, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_XMLUseXSLTWhenSaving(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1da, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_XMLSaveThroughXSLT()
    {
        CString result;
        InvokeHelper(0x1db, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_XMLSaveThroughXSLT(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x1db, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_DocumentLibraryVersions()
    {
        LPDISPATCH result;
        InvokeHelper(0x1dc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_ReadingModeLayoutFrozen()
    {
        BOOL result;
        InvokeHelper(0x1e1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ReadingModeLayoutFrozen(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1e1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_RemoveDateAndTime()
    {
        BOOL result;
        InvokeHelper(0x1e4, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_RemoveDateAndTime(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1e4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void SendFaxOverInternet(VARIANT * Recipients, VARIANT * Subject, VARIANT * ShowMessage)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1d0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Recipients, Subject, ShowMessage);
    }
    void TransformDocument(LPCTSTR Path, BOOL DataOnly)
    {
        static BYTE parms[] = VTS_BSTR VTS_BOOL;
        InvokeHelper(0x1f4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Path, DataOnly);
    }
    void Protect(long Type, VARIANT * NoReset, VARIANT * Password, VARIANT * UseIRM, VARIANT * EnforceStyleLock)
    {
        static BYTE parms[] = VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1d3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type, NoReset, Password, UseIRM, EnforceStyleLock);
    }
    void SelectAllEditableRanges(VARIANT * EditorID)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x1d4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, EditorID);
    }
    void DeleteAllEditableRanges(VARIANT * EditorID)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x1d5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, EditorID);
    }
    void DeleteAllInkAnnotations()
    {
        InvokeHelper(0x1df, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void AddDocumentWorkspaceHeader(BOOL RichFormat, LPCTSTR Url, LPCTSTR Title, LPCTSTR Description, LPCTSTR ID)
    {
        static BYTE parms[] = VTS_BOOL VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR;
        InvokeHelper(0x1e2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, RichFormat, Url, Title, Description, ID);
    }
    void RemoveDocumentWorkspaceHeader(LPCTSTR ID)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x1e3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ID);
    }
    void Compare(LPCTSTR Name, VARIANT * AuthorName, VARIANT * CompareTarget, VARIANT * DetectFormatChanges, VARIANT * IgnoreAllComparisonWarnings, VARIANT * AddToRecentFiles, VARIANT * RemovePersonalInformation, VARIANT * RemoveDateAndTime)
    {
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1e5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles, RemovePersonalInformation, RemoveDateAndTime);
    }
    void RemoveLockedStyles()
    {
        InvokeHelper(0x1e7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_ChildNodeSuggestions()
    {
        LPDISPATCH result;
        InvokeHelper(0x1e6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH SelectSingleNode(LPCTSTR XPath, LPCTSTR PrefixMapping, BOOL FastSearchSkippingTextNodes)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BOOL;
        InvokeHelper(0x1e8, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, XPath, PrefixMapping, FastSearchSkippingTextNodes);
        return result;
    }
    LPDISPATCH SelectNodes(LPCTSTR XPath, LPCTSTR PrefixMapping, BOOL FastSearchSkippingTextNodes)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BOOL;
        InvokeHelper(0x1e9, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, XPath, PrefixMapping, FastSearchSkippingTextNodes);
        return result;
    }
    LPDISPATCH get_XMLSchemaViolations()
    {
        LPDISPATCH result;
        InvokeHelper(0x1ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_ReadingLayoutSizeX()
    {
        long result;
        InvokeHelper(0x1eb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_ReadingLayoutSizeX(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1eb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_ReadingLayoutSizeY()
    {
        long result;
        InvokeHelper(0x1ec, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_ReadingLayoutSizeY(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1ec, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_StyleSortMethod()
    {
        long result;
        InvokeHelper(0x1ed, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_StyleSortMethod(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1ed, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_ContentTypeProperties()
    {
        LPDISPATCH result;
        InvokeHelper(0x1f0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_TrackMoves()
    {
        BOOL result;
        InvokeHelper(0x1f3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_TrackMoves(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1f3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_TrackFormatting()
    {
        BOOL result;
        InvokeHelper(0x1f6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_TrackFormatting(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x1f6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void get_Dummy1()
    {
        InvokeHelper(0x1f7, DISPATCH_PROPERTYGET, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_OMaths()
    {
        LPDISPATCH result;
        InvokeHelper(0x1f8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void RemoveDocumentInformation(long RemoveDocInfoType)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x1ef, DISPATCH_METHOD, VT_EMPTY, NULL, parms, RemoveDocInfoType);
    }
    void CheckInWithVersion(BOOL SaveChanges, VARIANT * Comments, BOOL MakePublic, VARIANT * VersionType)
    {
        static BYTE parms[] = VTS_BOOL VTS_PVARIANT VTS_BOOL VTS_PVARIANT;
        InvokeHelper(0x1f5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SaveChanges, Comments, MakePublic, VersionType);
    }
    void Dummy2()
    {
        InvokeHelper(0x1f9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void get_Dummy3()
    {
        InvokeHelper(0x1fa, DISPATCH_PROPERTYGET, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_ServerPolicy()
    {
        LPDISPATCH result;
        InvokeHelper(0x1fb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ContentControls()
    {
        LPDISPATCH result;
        InvokeHelper(0x1fc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_DocumentInspectors()
    {
        LPDISPATCH result;
        InvokeHelper(0x1fe, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void LockServerFile()
    {
        InvokeHelper(0x1fd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH GetWorkflowTasks()
    {
        LPDISPATCH result;
        InvokeHelper(0x1ff, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH GetWorkflowTemplates()
    {
        LPDISPATCH result;
        InvokeHelper(0x200, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void Dummy4()
    {
        InvokeHelper(0x202, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void AddMeetingWorkspaceHeader(BOOL SkipIfAbsent, LPCTSTR Url, LPCTSTR Title, LPCTSTR Description, LPCTSTR ID)
    {
        static BYTE parms[] = VTS_BOOL VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR;
        InvokeHelper(0x203, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SkipIfAbsent, Url, Title, Description, ID);
    }
    LPDISPATCH get_Bibliography()
    {
        LPDISPATCH result;
        InvokeHelper(0x204, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_LockTheme()
    {
        BOOL result;
        InvokeHelper(0x205, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_LockTheme(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x205, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_LockQuickStyleSet()
    {
        BOOL result;
        InvokeHelper(0x206, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_LockQuickStyleSet(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x206, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_OriginalDocumentTitle()
    {
        CString result;
        InvokeHelper(0x207, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    CString get_RevisedDocumentTitle()
    {
        CString result;
        InvokeHelper(0x208, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_CustomXMLParts()
    {
        LPDISPATCH result;
        InvokeHelper(0x209, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_FormattingShowNextLevel()
    {
        BOOL result;
        InvokeHelper(0x20a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_FormattingShowNextLevel(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x20a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_FormattingShowUserStyleName()
    {
        BOOL result;
        InvokeHelper(0x20b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_FormattingShowUserStyleName(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x20b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void SaveAsQuickStyleSet(LPCTSTR FileName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x20c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
    }
    void ApplyQuickStyleSet(LPCTSTR Name)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x20d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name);
    }
    LPDISPATCH get_Research()
    {
        LPDISPATCH result;
        InvokeHelper(0x20e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_Final()
    {
        BOOL result;
        InvokeHelper(0x20f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_Final(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x20f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_OMathBreakBin()
    {
        long result;
        InvokeHelper(0x210, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_OMathBreakBin(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x210, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_OMathBreakSub()
    {
        long result;
        InvokeHelper(0x211, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_OMathBreakSub(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x211, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_OMathJc()
    {
        long result;
        InvokeHelper(0x212, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_OMathJc(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x212, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_OMathLeftMargin()
    {
        float result;
        InvokeHelper(0x213, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_OMathLeftMargin(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x213, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_OMathRightMargin()
    {
        float result;
        InvokeHelper(0x214, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_OMathRightMargin(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x214, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_OMathWrap()
    {
        float result;
        InvokeHelper(0x217, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_OMathWrap(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x217, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_OMathIntSubSupLim()
    {
        BOOL result;
        InvokeHelper(0x218, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_OMathIntSubSupLim(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x218, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_OMathNarySupSubLim()
    {
        BOOL result;
        InvokeHelper(0x219, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_OMathNarySupSubLim(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x219, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_OMathSmallFrac()
    {
        BOOL result;
        InvokeHelper(0x21b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_OMathSmallFrac(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x21b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_WordOpenXML()
    {
        CString result;
        InvokeHelper(0x21e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_DocumentTheme()
    {
        LPDISPATCH result;
        InvokeHelper(0x221, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ApplyDocumentTheme(LPCTSTR FileName)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x222, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
    }
    BOOL get_HasVBProject()
    {
        BOOL result;
        InvokeHelper(0x224, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH SelectLinkedControls(LPDISPATCH Node)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x225, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Node);
        return result;
    }
    LPDISPATCH SelectUnlinkedControls(LPDISPATCH Stream)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x226, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Stream);
        return result;
    }
    LPDISPATCH SelectContentControlsByTitle(LPCTSTR Title)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x227, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Title);
        return result;
    }
    void ExportAsFixedFormat(LPCTSTR OutputFileName, long ExportFormat, BOOL OpenAfterExport, long OptimizeFor, long Range, long From, long To, long Item, BOOL IncludeDocProps, BOOL KeepIRM, long CreateBookmarks, BOOL DocStructureTags, BOOL BitmapMissingFonts, BOOL UseISO19005_1, VARIANT * FixedFormatExtClassPtr)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_BOOL VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BOOL VTS_BOOL VTS_I4 VTS_BOOL VTS_BOOL VTS_BOOL VTS_PVARIANT;
        InvokeHelper(0x228, DISPATCH_METHOD, VT_EMPTY, NULL, parms, OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr);
    }
    void FreezeLayout()
    {
        InvokeHelper(0x229, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void UnfreezeLayout()
    {
        InvokeHelper(0x22a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    CString get_OMathFontName()
    {
        CString result;
        InvokeHelper(0x22b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_OMathFontName(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x22b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void DowngradeDocument()
    {
        InvokeHelper(0x22e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    CString get_EncryptionProvider()
    {
        CString result;
        InvokeHelper(0x22f, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_EncryptionProvider(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x22f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_UseMathDefaults()
    {
        BOOL result;
        InvokeHelper(0x230, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_UseMathDefaults(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x230, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_CurrentRsid()
    {
        long result;
        InvokeHelper(0x233, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void Convert()
    {
        InvokeHelper(0x231, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH SelectContentControlsByTag(LPCTSTR Tag)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x232, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Tag);
        return result;
    }
    void ConvertAutoHyphens()
    {
        InvokeHelper(0x28a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    long get_DocID()
    {
        long result;
        InvokeHelper(0x234, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void ApplyQuickStyleSet2(VARIANT * Style)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x236, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Style);
    }
    long get_CompatibilityMode()
    {
        long result;
        InvokeHelper(0x237, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void SaveAs2(VARIANT * FileName, VARIANT * FileFormat, VARIANT * LockComments, VARIANT * Password, VARIANT * AddToRecentFiles, VARIANT * WritePassword, VARIANT * ReadOnlyRecommended, VARIANT * EmbedTrueTypeFonts, VARIANT * SaveNativePictureFormat, VARIANT * SaveFormsData, VARIANT * SaveAsAOCELetter, VARIANT * Encoding, VARIANT * InsertLineBreaks, VARIANT * AllowSubstitutions, VARIANT * LineEnding, VARIANT * AddBiDiMarks, VARIANT * CompatibilityMode)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x238, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks, CompatibilityMode);
    }
    LPDISPATCH get_CoAuthoring()
    {
        LPDISPATCH result;
        InvokeHelper(0x258, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void SetCompatibilityMode(long Mode)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x23b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Mode);
    }
    LPDISPATCH get_Broadcast()
    {
        LPDISPATCH result;
        InvokeHelper(0x262, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long ReturnToLastReadPosition()
    {
        long result;
        InvokeHelper(0x23c, DISPATCH_METHOD, VT_I4, (void*)&result, NULL);
        return result;
    }
    BOOL get_ChartDataPointTrack()
    {
        BOOL result;
        InvokeHelper(0x26a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ChartDataPointTrack(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x26a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_IsInAutosave()
    {
        BOOL result;
        InvokeHelper(0x26b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void SaveCopyAs(VARIANT * FileName, VARIANT * FileFormat, VARIANT * LockComments, VARIANT * Password, VARIANT * AddToRecentFiles, VARIANT * WritePassword, VARIANT * ReadOnlyRecommended, VARIANT * EmbedTrueTypeFonts, VARIANT * SaveNativePictureFormat, VARIANT * SaveFormsData, VARIANT * SaveAsAOCELetter, VARIANT * Encoding, VARIANT * InsertLineBreaks, VARIANT * AllowSubstitutions, VARIANT * LineEnding, VARIANT * AddBiDiMarks, VARIANT * CompatibilityMode)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x26d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks, CompatibilityMode);
    }

    // _Document 属性
public:

};

// CSelection 包装器类

class CSelection : public COleDispatchDriver
{
public:
    CSelection(){} // 调用 COleDispatchDriver 默认构造函数
    CSelection(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
    CSelection(const CSelection& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

    // 特性
public:

    // 操作
public:


    // Selection 方法
public:
    CString get_Text()
    {
        CString result;
        InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_Text(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_FormattedText()
    {
        LPDISPATCH result;
        InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_FormattedText(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Start()
    {
        long result;
        InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Start(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_End()
    {
        long result;
        InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_End(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Font()
    {
        LPDISPATCH result;
        InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_Font(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Type()
    {
        long result;
        InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_StoryType()
    {
        long result;
        InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    VARIANT get_Style()
    {
        VARIANT result;
        InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    void put_Style(VARIANT * newValue)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Tables()
    {
        LPDISPATCH result;
        InvokeHelper(0x32, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Words()
    {
        LPDISPATCH result;
        InvokeHelper(0x33, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Sentences()
    {
        LPDISPATCH result;
        InvokeHelper(0x34, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Characters()
    {
        LPDISPATCH result;
        InvokeHelper(0x35, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Footnotes()
    {
        LPDISPATCH result;
        InvokeHelper(0x36, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Endnotes()
    {
        LPDISPATCH result;
        InvokeHelper(0x37, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Comments()
    {
        LPDISPATCH result;
        InvokeHelper(0x38, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Cells()
    {
        LPDISPATCH result;
        InvokeHelper(0x39, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Sections()
    {
        LPDISPATCH result;
        InvokeHelper(0x3a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Paragraphs()
    {
        LPDISPATCH result;
        InvokeHelper(0x3b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Borders()
    {
        LPDISPATCH result;
        InvokeHelper(0x44c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_Borders(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x44c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Shading()
    {
        LPDISPATCH result;
        InvokeHelper(0x3d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Fields()
    {
        LPDISPATCH result;
        InvokeHelper(0x40, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FormFields()
    {
        LPDISPATCH result;
        InvokeHelper(0x41, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Frames()
    {
        LPDISPATCH result;
        InvokeHelper(0x42, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ParagraphFormat()
    {
        LPDISPATCH result;
        InvokeHelper(0x44e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_ParagraphFormat(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x44e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_PageSetup()
    {
        LPDISPATCH result;
        InvokeHelper(0x44d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_PageSetup(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x44d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Bookmarks()
    {
        LPDISPATCH result;
        InvokeHelper(0x4b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_StoryLength()
    {
        long result;
        InvokeHelper(0x98, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_LanguageID()
    {
        long result;
        InvokeHelper(0x99, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_LanguageID(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x99, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_LanguageIDFarEast()
    {
        long result;
        InvokeHelper(0x9a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_LanguageIDFarEast(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x9a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_LanguageIDOther()
    {
        long result;
        InvokeHelper(0x9b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_LanguageIDOther(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x9b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Hyperlinks()
    {
        LPDISPATCH result;
        InvokeHelper(0x9c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Columns()
    {
        LPDISPATCH result;
        InvokeHelper(0x12e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Rows()
    {
        LPDISPATCH result;
        InvokeHelper(0x12f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_HeaderFooter()
    {
        LPDISPATCH result;
        InvokeHelper(0x132, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_IsEndOfRowMark()
    {
        BOOL result;
        InvokeHelper(0x133, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    long get_BookmarkID()
    {
        long result;
        InvokeHelper(0x134, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_PreviousBookmarkID()
    {
        long result;
        InvokeHelper(0x135, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Find()
    {
        LPDISPATCH result;
        InvokeHelper(0x106, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Range()
    {
        LPDISPATCH result;
        InvokeHelper(0x190, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    VARIANT get_Information(long Type)
    {
        VARIANT result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x191, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms, Type);
        return result;
    }
    long get_Flags()
    {
        long result;
        InvokeHelper(0x192, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Flags(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x192, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_Active()
    {
        BOOL result;
        InvokeHelper(0x193, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    BOOL get_StartIsActive()
    {
        BOOL result;
        InvokeHelper(0x194, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_StartIsActive(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x194, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_IPAtEndOfLine()
    {
        BOOL result;
        InvokeHelper(0x195, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    BOOL get_ExtendMode()
    {
        BOOL result;
        InvokeHelper(0x196, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ExtendMode(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x196, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_ColumnSelectMode()
    {
        BOOL result;
        InvokeHelper(0x197, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ColumnSelectMode(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x197, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Orientation()
    {
        long result;
        InvokeHelper(0x19a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Orientation(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x19a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_InlineShapes()
    {
        LPDISPATCH result;
        InvokeHelper(0x19b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Application()
    {
        LPDISPATCH result;
        InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Creator()
    {
        long result;
        InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Parent()
    {
        LPDISPATCH result;
        InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Document()
    {
        LPDISPATCH result;
        InvokeHelper(0x3eb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ShapeRange()
    {
        LPDISPATCH result;
        InvokeHelper(0x3ec, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void Select()
    {
        InvokeHelper(0xffff, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SetRange(long Start, long End)
    {
        static BYTE parms[] = VTS_I4 VTS_I4;
        InvokeHelper(0x64, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Start, End);
    }
    void Collapse(VARIANT * Direction)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Direction);
    }
    void InsertBefore(LPCTSTR Text)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text);
    }
    void InsertAfter(LPCTSTR Text)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text);
    }
    LPDISPATCH Next(VARIANT * Unit, VARIANT * Count)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x69, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Unit, Count);
        return result;
    }
    LPDISPATCH Previous(VARIANT * Unit, VARIANT * Count)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Unit, Count);
        return result;
    }
    long StartOf(VARIANT * Unit, VARIANT * Extend)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6b, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Extend);
        return result;
    }
    long EndOf(VARIANT * Unit, VARIANT * Extend)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6c, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Extend);
        return result;
    }
    long Move(VARIANT * Unit, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6d, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
        return result;
    }
    long MoveStart(VARIANT * Unit, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6e, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
        return result;
    }
    long MoveEnd(VARIANT * Unit, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6f, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
        return result;
    }
    long MoveWhile(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x70, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    long MoveStartWhile(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x71, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    long MoveEndWhile(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x72, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    long MoveUntil(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x73, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    long MoveStartUntil(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x74, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    long MoveEndUntil(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x75, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    void Cut()
    {
        InvokeHelper(0x77, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Copy()
    {
        InvokeHelper(0x78, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Paste()
    {
        InvokeHelper(0x79, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertBreak(VARIANT * Type)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x7a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type);
    }
    void InsertFile(LPCTSTR FileName, VARIANT * Range, VARIANT * ConfirmConversions, VARIANT * Link, VARIANT * Attachment)
    {
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x7b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, Range, ConfirmConversions, Link, Attachment);
    }
    BOOL InStory(LPDISPATCH Range)
    {
        BOOL result;
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x7d, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Range);
        return result;
    }
    BOOL InRange(LPDISPATCH Range)
    {
        BOOL result;
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x7e, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Range);
        return result;
    }
    long Delete(VARIANT * Unit, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x7f, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
        return result;
    }
    long Expand(VARIANT * Unit)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x81, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit);
        return result;
    }
    void InsertParagraph()
    {
        InvokeHelper(0xa0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertParagraphAfter()
    {
        InvokeHelper(0xa1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH ConvertToTableOld(VARIANT * Separator, VARIANT * NumRows, VARIANT * NumColumns, VARIANT * InitialColumnWidth, VARIANT * Format, VARIANT * ApplyBorders, VARIANT * ApplyShading, VARIANT * ApplyFont, VARIANT * ApplyColor, VARIANT * ApplyHeadingRows, VARIANT * ApplyLastRow, VARIANT * ApplyFirstColumn, VARIANT * ApplyLastColumn, VARIANT * AutoFit)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit);
        return result;
    }
    void InsertDateTimeOld(VARIANT * DateTimeFormat, VARIANT * InsertAsField, VARIANT * InsertAsFullWidth)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DateTimeFormat, InsertAsField, InsertAsFullWidth);
    }
    void InsertSymbol(long CharacterNumber, VARIANT * Font, VARIANT * Unicode, VARIANT * Bias)
    {
        static BYTE parms[] = VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, CharacterNumber, Font, Unicode, Bias);
    }
    void InsertCrossReference_2002(VARIANT * ReferenceType, long ReferenceKind, VARIANT * ReferenceItem, VARIANT * InsertAsHyperlink, VARIANT * IncludePosition)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition);
    }
    void InsertCaptionXP(VARIANT * Label, VARIANT * Title, VARIANT * TitleAutoText, VARIANT * Position)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Label, Title, TitleAutoText, Position);
    }
    void CopyAsPicture()
    {
        InvokeHelper(0xa7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SortOld(VARIANT * ExcludeHeader, VARIANT * FieldNumber, VARIANT * SortFieldType, VARIANT * SortOrder, VARIANT * FieldNumber2, VARIANT * SortFieldType2, VARIANT * SortOrder2, VARIANT * FieldNumber3, VARIANT * SortFieldType3, VARIANT * SortOrder3, VARIANT * SortColumn, VARIANT * Separator, VARIANT * CaseSensitive, VARIANT * LanguageID)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, LanguageID);
    }
    void SortAscending()
    {
        InvokeHelper(0xa9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SortDescending()
    {
        InvokeHelper(0xaa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    BOOL IsEqual(LPDISPATCH Range)
    {
        BOOL result;
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0xab, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Range);
        return result;
    }
    float Calculate()
    {
        float result;
        InvokeHelper(0xac, DISPATCH_METHOD, VT_R4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH GoTo(VARIANT * What, VARIANT * Which, VARIANT * Count, VARIANT * Name)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xad, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, What, Which, Count, Name);
        return result;
    }
    LPDISPATCH GoToNext(long What)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0xae, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, What);
        return result;
    }
    LPDISPATCH GoToPrevious(long What)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0xaf, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, What);
        return result;
    }
    void PasteSpecial(VARIANT * IconIndex, VARIANT * Link, VARIANT * Placement, VARIANT * DisplayAsIcon, VARIANT * DataType, VARIANT * IconFileName, VARIANT * IconLabel)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xb0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel);
    }
    LPDISPATCH PreviousField()
    {
        LPDISPATCH result;
        InvokeHelper(0xb1, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH NextField()
    {
        LPDISPATCH result;
        InvokeHelper(0xb2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void InsertParagraphBefore()
    {
        InvokeHelper(0xd4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertCells(VARIANT * ShiftCells)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0xd6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ShiftCells);
    }
    void Extend(VARIANT * Character)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x12c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Character);
    }
    void Shrink()
    {
        InvokeHelper(0x12d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    long MoveLeft(VARIANT * Unit, VARIANT * Count, VARIANT * Extend)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1f4, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count, Extend);
        return result;
    }
    long MoveRight(VARIANT * Unit, VARIANT * Count, VARIANT * Extend)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1f5, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count, Extend);
        return result;
    }
    long MoveUp(VARIANT * Unit, VARIANT * Count, VARIANT * Extend)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1f6, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count, Extend);
        return result;
    }
    long MoveDown(VARIANT * Unit, VARIANT * Count, VARIANT * Extend)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1f7, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count, Extend);
        return result;
    }
    long HomeKey(VARIANT * Unit, VARIANT * Extend)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1f8, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Extend);
        return result;
    }
    long EndKey(VARIANT * Unit, VARIANT * Extend)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1f9, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Extend);
        return result;
    }
    void EscapeKey()
    {
        InvokeHelper(0x1fa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void TypeText(LPCTSTR Text)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x1fb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text);
    }
    void CopyFormat()
    {
        InvokeHelper(0x1fd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void PasteFormat()
    {
        InvokeHelper(0x1fe, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void TypeParagraph()
    {
        InvokeHelper(0x200, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void TypeBackspace()
    {
        InvokeHelper(0x201, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void NextSubdocument()
    {
        InvokeHelper(0x202, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void PreviousSubdocument()
    {
        InvokeHelper(0x203, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SelectColumn()
    {
        InvokeHelper(0x204, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SelectCurrentFont()
    {
        InvokeHelper(0x205, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SelectCurrentAlignment()
    {
        InvokeHelper(0x206, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SelectCurrentSpacing()
    {
        InvokeHelper(0x207, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SelectCurrentIndent()
    {
        InvokeHelper(0x208, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SelectCurrentTabs()
    {
        InvokeHelper(0x209, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SelectCurrentColor()
    {
        InvokeHelper(0x20a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void CreateTextbox()
    {
        InvokeHelper(0x20b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void WholeStory()
    {
        InvokeHelper(0x20c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SelectRow()
    {
        InvokeHelper(0x20d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SplitTable()
    {
        InvokeHelper(0x20e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertRows(VARIANT * NumRows)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x210, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumRows);
    }
    void InsertColumns()
    {
        InvokeHelper(0x211, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertFormula(VARIANT * Formula, VARIANT * NumberFormat)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x212, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Formula, NumberFormat);
    }
    LPDISPATCH NextRevision(VARIANT * Wrap)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x213, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Wrap);
        return result;
    }
    LPDISPATCH PreviousRevision(VARIANT * Wrap)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x214, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Wrap);
        return result;
    }
    void PasteAsNestedTable()
    {
        InvokeHelper(0x215, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH CreateAutoTextEntry(LPCTSTR Name, LPCTSTR StyleName)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_BSTR VTS_BSTR;
        InvokeHelper(0x216, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, StyleName);
        return result;
    }
    void DetectLanguage()
    {
        InvokeHelper(0x217, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SelectCell()
    {
        InvokeHelper(0x218, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertRowsBelow(VARIANT * NumRows)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x219, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumRows);
    }
    void InsertColumnsRight()
    {
        InvokeHelper(0x21a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertRowsAbove(VARIANT * NumRows)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x21b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumRows);
    }
    void RtlRun()
    {
        InvokeHelper(0x258, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void LtrRun()
    {
        InvokeHelper(0x259, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void BoldRun()
    {
        InvokeHelper(0x25a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ItalicRun()
    {
        InvokeHelper(0x25b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void RtlPara()
    {
        InvokeHelper(0x25d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void LtrPara()
    {
        InvokeHelper(0x25e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertDateTime(VARIANT * DateTimeFormat, VARIANT * InsertAsField, VARIANT * InsertAsFullWidth, VARIANT * DateLanguage, VARIANT * CalendarType)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1bc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType);
    }
    void Sort2000(VARIANT * ExcludeHeader, VARIANT * FieldNumber, VARIANT * SortFieldType, VARIANT * SortOrder, VARIANT * FieldNumber2, VARIANT * SortFieldType2, VARIANT * SortOrder2, VARIANT * FieldNumber3, VARIANT * SortFieldType3, VARIANT * SortOrder3, VARIANT * SortColumn, VARIANT * Separator, VARIANT * CaseSensitive, VARIANT * BidiSort, VARIANT * IgnoreThe, VARIANT * IgnoreKashida, VARIANT * IgnoreDiacritics, VARIANT * IgnoreHe, VARIANT * LanguageID)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1bd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID);
    }
    LPDISPATCH ConvertToTable(VARIANT * Separator, VARIANT * NumRows, VARIANT * NumColumns, VARIANT * InitialColumnWidth, VARIANT * Format, VARIANT * ApplyBorders, VARIANT * ApplyShading, VARIANT * ApplyFont, VARIANT * ApplyColor, VARIANT * ApplyHeadingRows, VARIANT * ApplyLastRow, VARIANT * ApplyFirstColumn, VARIANT * ApplyLastColumn, VARIANT * AutoFit, VARIANT * AutoFitBehavior, VARIANT * DefaultTableBehavior)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1c9, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior);
        return result;
    }
    long get_NoProofing()
    {
        long result;
        InvokeHelper(0x3ed, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_NoProofing(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x3ed, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_TopLevelTables()
    {
        LPDISPATCH result;
        InvokeHelper(0x3ee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_LanguageDetected()
    {
        BOOL result;
        InvokeHelper(0x3ef, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_LanguageDetected(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x3ef, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_FitTextWidth()
    {
        float result;
        InvokeHelper(0x3f0, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_FitTextWidth(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x3f0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void ClearFormatting()
    {
        InvokeHelper(0x3f1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void PasteAppendTable()
    {
        InvokeHelper(0x3f2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_HTMLDivisions()
    {
        LPDISPATCH result;
        InvokeHelper(0x3f3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SmartTags()
    {
        LPDISPATCH result;
        InvokeHelper(0x3f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ChildShapeRange()
    {
        LPDISPATCH result;
        InvokeHelper(0x3fd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_HasChildShapeRange()
    {
        BOOL result;
        InvokeHelper(0x3fe, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FootnoteOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x400, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_EndnoteOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x401, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ToggleCharacterCode()
    {
        InvokeHelper(0x3f4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void PasteAndFormat(long Type)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x3f5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type);
    }
    void PasteExcelTable(BOOL LinkedToExcel, BOOL WordFormatting, BOOL RTF)
    {
        static BYTE parms[] = VTS_BOOL VTS_BOOL VTS_BOOL;
        InvokeHelper(0x3f6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, LinkedToExcel, WordFormatting, RTF);
    }
    void ShrinkDiscontiguousSelection()
    {
        InvokeHelper(0x3fb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertStyleSeparator()
    {
        InvokeHelper(0x3fc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Sort(VARIANT * ExcludeHeader, VARIANT * FieldNumber, VARIANT * SortFieldType, VARIANT * SortOrder, VARIANT * FieldNumber2, VARIANT * SortFieldType2, VARIANT * SortOrder2, VARIANT * FieldNumber3, VARIANT * SortFieldType3, VARIANT * SortOrder3, VARIANT * SortColumn, VARIANT * Separator, VARIANT * CaseSensitive, VARIANT * BidiSort, VARIANT * IgnoreThe, VARIANT * IgnoreKashida, VARIANT * IgnoreDiacritics, VARIANT * IgnoreHe, VARIANT * LanguageID, VARIANT * SubFieldNumber, VARIANT * SubFieldNumber2, VARIANT * SubFieldNumber3)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x3ff, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID, SubFieldNumber, SubFieldNumber2, SubFieldNumber3);
    }
    LPDISPATCH get_XMLNodes()
    {
        LPDISPATCH result;
        InvokeHelper(0x136, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_XMLParentNode()
    {
        LPDISPATCH result;
        InvokeHelper(0x137, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Editors()
    {
        LPDISPATCH result;
        InvokeHelper(0x139, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_XML(BOOL DataOnly)
    {
        CString result;
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x13a, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, DataOnly);
        return result;
    }
    VARIANT get_EnhMetaFileBits()
    {
        VARIANT result;
        InvokeHelper(0x13b, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH GoToEditableRange(VARIANT * EditorID)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x403, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, EditorID);
        return result;
    }
    void InsertXML(LPCTSTR XML, VARIANT * Transform)
    {
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT;
        InvokeHelper(0x404, DISPATCH_METHOD, VT_EMPTY, NULL, parms, XML, Transform);
    }
    void InsertCaption(VARIANT * Label, VARIANT * Title, VARIANT * TitleAutoText, VARIANT * Position, VARIANT * ExcludeLabel)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1a1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Label, Title, TitleAutoText, Position, ExcludeLabel);
    }
    void InsertCrossReference(VARIANT * ReferenceType, long ReferenceKind, VARIANT * ReferenceItem, VARIANT * InsertAsHyperlink, VARIANT * IncludePosition, VARIANT * SeparateNumbers, VARIANT * SeparatorString)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1a2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString);
    }
    LPDISPATCH get_OMaths()
    {
        LPDISPATCH result;
        InvokeHelper(0x13c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_WordOpenXML()
    {
        CString result;
        InvokeHelper(0x13d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void ClearParagraphStyle()
    {
        InvokeHelper(0x406, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ClearCharacterAllFormatting()
    {
        InvokeHelper(0x407, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ClearCharacterStyle()
    {
        InvokeHelper(0x408, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ClearCharacterDirectFormatting()
    {
        InvokeHelper(0x409, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_ContentControls()
    {
        LPDISPATCH result;
        InvokeHelper(0x40a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ParentContentControl()
    {
        LPDISPATCH result;
        InvokeHelper(0x40b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ExportAsFixedFormat(LPCTSTR OutputFileName, long ExportFormat, BOOL OpenAfterExport, long OptimizeFor, BOOL ExportCurrentPage, long Item, BOOL IncludeDocProps, BOOL KeepIRM, long CreateBookmarks, BOOL DocStructureTags, BOOL BitmapMissingFonts, BOOL UseISO19005_1, VARIANT * FixedFormatExtClassPtr)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_BOOL VTS_I4 VTS_BOOL VTS_I4 VTS_BOOL VTS_BOOL VTS_I4 VTS_BOOL VTS_BOOL VTS_BOOL VTS_PVARIANT;
        InvokeHelper(0x40c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr);
    }
    void ReadingModeGrowFont()
    {
        InvokeHelper(0x40d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ReadingModeShrinkFont()
    {
        InvokeHelper(0x40e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ClearParagraphAllFormatting()
    {
        InvokeHelper(0x40f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ClearParagraphDirectFormatting()
    {
        InvokeHelper(0x410, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertNewPage()
    {
        InvokeHelper(0x411, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SortByHeadings(VARIANT * SortFieldType, VARIANT * SortOrder, VARIANT * CaseSensitive, VARIANT * BidiSort, VARIANT * IgnoreThe, VARIANT * IgnoreKashida, VARIANT * IgnoreDiacritics, VARIANT * IgnoreHe, VARIANT * LanguageID)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x412, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SortFieldType, SortOrder, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID);
    }

    // Selection 属性
public:

};


// CRange 包装器类

class CRange : public COleDispatchDriver
{
public:
    CRange(){} // 调用 COleDispatchDriver 默认构造函数
    CRange(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
    CRange(const CRange& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

    // 特性
public:

    // 操作
public:


    // Range 方法
public:
    CString get_Text()
    {
        CString result;
        InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_Text(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_FormattedText()
    {
        LPDISPATCH result;
        InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_FormattedText(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Start()
    {
        long result;
        InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Start(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_End()
    {
        long result;
        InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_End(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Font()
    {
        LPDISPATCH result;
        InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_Font(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Duplicate()
    {
        LPDISPATCH result;
        InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_StoryType()
    {
        long result;
        InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Tables()
    {
        LPDISPATCH result;
        InvokeHelper(0x32, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Words()
    {
        LPDISPATCH result;
        InvokeHelper(0x33, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Sentences()
    {
        LPDISPATCH result;
        InvokeHelper(0x34, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Characters()
    {
        LPDISPATCH result;
        InvokeHelper(0x35, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Footnotes()
    {
        LPDISPATCH result;
        InvokeHelper(0x36, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Endnotes()
    {
        LPDISPATCH result;
        InvokeHelper(0x37, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Comments()
    {
        LPDISPATCH result;
        InvokeHelper(0x38, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Cells()
    {
        LPDISPATCH result;
        InvokeHelper(0x39, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Sections()
    {
        LPDISPATCH result;
        InvokeHelper(0x3a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Paragraphs()
    {
        LPDISPATCH result;
        InvokeHelper(0x3b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Borders()
    {
        LPDISPATCH result;
        InvokeHelper(0x44c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_Borders(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x44c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Shading()
    {
        LPDISPATCH result;
        InvokeHelper(0x3d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_TextRetrievalMode()
    {
        LPDISPATCH result;
        InvokeHelper(0x3e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_TextRetrievalMode(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x3e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Fields()
    {
        LPDISPATCH result;
        InvokeHelper(0x40, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FormFields()
    {
        LPDISPATCH result;
        InvokeHelper(0x41, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Frames()
    {
        LPDISPATCH result;
        InvokeHelper(0x42, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ParagraphFormat()
    {
        LPDISPATCH result;
        InvokeHelper(0x44e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_ParagraphFormat(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x44e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_ListFormat()
    {
        LPDISPATCH result;
        InvokeHelper(0x44, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Bookmarks()
    {
        LPDISPATCH result;
        InvokeHelper(0x4b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Application()
    {
        LPDISPATCH result;
        InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Creator()
    {
        long result;
        InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Parent()
    {
        LPDISPATCH result;
        InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Bold()
    {
        long result;
        InvokeHelper(0x82, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Bold(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x82, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Italic()
    {
        long result;
        InvokeHelper(0x83, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Italic(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x83, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Underline()
    {
        long result;
        InvokeHelper(0x8b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Underline(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x8b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_EmphasisMark()
    {
        long result;
        InvokeHelper(0x8c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_EmphasisMark(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x8c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_DisableCharacterSpaceGrid()
    {
        BOOL result;
        InvokeHelper(0x8d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_DisableCharacterSpaceGrid(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x8d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Revisions()
    {
        LPDISPATCH result;
        InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    VARIANT get_Style()
    {
        VARIANT result;
        InvokeHelper(0x97, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    void put_Style(VARIANT * newValue)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x97, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_StoryLength()
    {
        long result;
        InvokeHelper(0x98, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_LanguageID()
    {
        long result;
        InvokeHelper(0x99, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_LanguageID(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x99, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_SynonymInfo()
    {
        LPDISPATCH result;
        InvokeHelper(0x9b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Hyperlinks()
    {
        LPDISPATCH result;
        InvokeHelper(0x9c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ListParagraphs()
    {
        LPDISPATCH result;
        InvokeHelper(0x9d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Subdocuments()
    {
        LPDISPATCH result;
        InvokeHelper(0x9f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_GrammarChecked()
    {
        BOOL result;
        InvokeHelper(0x104, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_GrammarChecked(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x104, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_SpellingChecked()
    {
        BOOL result;
        InvokeHelper(0x105, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_SpellingChecked(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x105, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_HighlightColorIndex()
    {
        long result;
        InvokeHelper(0x12d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_HighlightColorIndex(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x12d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Columns()
    {
        LPDISPATCH result;
        InvokeHelper(0x12e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Rows()
    {
        LPDISPATCH result;
        InvokeHelper(0x12f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_CanEdit()
    {
        long result;
        InvokeHelper(0x130, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_CanPaste()
    {
        long result;
        InvokeHelper(0x131, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    BOOL get_IsEndOfRowMark()
    {
        BOOL result;
        InvokeHelper(0x133, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    long get_BookmarkID()
    {
        long result;
        InvokeHelper(0x134, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    long get_PreviousBookmarkID()
    {
        long result;
        InvokeHelper(0x135, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Find()
    {
        LPDISPATCH result;
        InvokeHelper(0x106, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_PageSetup()
    {
        LPDISPATCH result;
        InvokeHelper(0x44d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void put_PageSetup(LPDISPATCH newValue)
    {
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x44d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_ShapeRange()
    {
        LPDISPATCH result;
        InvokeHelper(0x137, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Case()
    {
        long result;
        InvokeHelper(0x138, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Case(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x138, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    VARIANT get_Information(long Type)
    {
        VARIANT result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x139, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms, Type);
        return result;
    }
    LPDISPATCH get_ReadabilityStatistics()
    {
        LPDISPATCH result;
        InvokeHelper(0x13a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_GrammaticalErrors()
    {
        LPDISPATCH result;
        InvokeHelper(0x13b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SpellingErrors()
    {
        LPDISPATCH result;
        InvokeHelper(0x13c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_Orientation()
    {
        long result;
        InvokeHelper(0x13d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Orientation(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x13d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_InlineShapes()
    {
        LPDISPATCH result;
        InvokeHelper(0x13f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_NextStoryRange()
    {
        LPDISPATCH result;
        InvokeHelper(0x140, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_LanguageIDFarEast()
    {
        long result;
        InvokeHelper(0x141, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_LanguageIDFarEast(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x141, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_LanguageIDOther()
    {
        long result;
        InvokeHelper(0x142, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_LanguageIDOther(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x142, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    void Select()
    {
        InvokeHelper(0xffff, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SetRange(long Start, long End)
    {
        static BYTE parms[] = VTS_I4 VTS_I4;
        InvokeHelper(0x64, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Start, End);
    }
    void Collapse(VARIANT * Direction)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Direction);
    }
    void InsertBefore(LPCTSTR Text)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text);
    }
    void InsertAfter(LPCTSTR Text)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text);
    }
    LPDISPATCH Next(VARIANT * Unit, VARIANT * Count)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x69, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Unit, Count);
        return result;
    }
    LPDISPATCH Previous(VARIANT * Unit, VARIANT * Count)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Unit, Count);
        return result;
    }
    long StartOf(VARIANT * Unit, VARIANT * Extend)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6b, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Extend);
        return result;
    }
    long EndOf(VARIANT * Unit, VARIANT * Extend)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6c, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Extend);
        return result;
    }
    long Move(VARIANT * Unit, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6d, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
        return result;
    }
    long MoveStart(VARIANT * Unit, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6e, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
        return result;
    }
    long MoveEnd(VARIANT * Unit, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x6f, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
        return result;
    }
    long MoveWhile(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x70, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    long MoveStartWhile(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x71, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    long MoveEndWhile(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x72, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    long MoveUntil(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x73, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    long MoveStartUntil(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x74, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    long MoveEndUntil(VARIANT * Cset, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x75, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Cset, Count);
        return result;
    }
    void Cut()
    {
        InvokeHelper(0x77, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Copy()
    {
        InvokeHelper(0x78, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void Paste()
    {
        InvokeHelper(0x79, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertBreak(VARIANT * Type)
    {
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x7a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type);
    }
    void InsertFile(LPCTSTR FileName, VARIANT * Range, VARIANT * ConfirmConversions, VARIANT * Link, VARIANT * Attachment)
    {
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x7b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, Range, ConfirmConversions, Link, Attachment);
    }
    BOOL InStory(LPDISPATCH Range)
    {
        BOOL result;
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x7d, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Range);
        return result;
    }
    BOOL InRange(LPDISPATCH Range)
    {
        BOOL result;
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0x7e, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Range);
        return result;
    }
    long Delete(VARIANT * Unit, VARIANT * Count)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x7f, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit, Count);
        return result;
    }
    void WholeStory()
    {
        InvokeHelper(0x80, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    long Expand(VARIANT * Unit)
    {
        long result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x81, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Unit);
        return result;
    }
    void InsertParagraph()
    {
        InvokeHelper(0xa0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertParagraphAfter()
    {
        InvokeHelper(0xa1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH ConvertToTableOld(VARIANT * Separator, VARIANT * NumRows, VARIANT * NumColumns, VARIANT * InitialColumnWidth, VARIANT * Format, VARIANT * ApplyBorders, VARIANT * ApplyShading, VARIANT * ApplyFont, VARIANT * ApplyColor, VARIANT * ApplyHeadingRows, VARIANT * ApplyLastRow, VARIANT * ApplyFirstColumn, VARIANT * ApplyLastColumn, VARIANT * AutoFit)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit);
        return result;
    }
    void InsertDateTimeOld(VARIANT * DateTimeFormat, VARIANT * InsertAsField, VARIANT * InsertAsFullWidth)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DateTimeFormat, InsertAsField, InsertAsFullWidth);
    }
    void InsertSymbol(long CharacterNumber, VARIANT * Font, VARIANT * Unicode, VARIANT * Bias)
    {
        static BYTE parms[] = VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, CharacterNumber, Font, Unicode, Bias);
    }
    void InsertCrossReference_2002(VARIANT * ReferenceType, long ReferenceKind, VARIANT * ReferenceItem, VARIANT * InsertAsHyperlink, VARIANT * IncludePosition)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition);
    }
    void InsertCaptionXP(VARIANT * Label, VARIANT * Title, VARIANT * TitleAutoText, VARIANT * Position)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Label, Title, TitleAutoText, Position);
    }
    void CopyAsPicture()
    {
        InvokeHelper(0xa7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SortOld(VARIANT * ExcludeHeader, VARIANT * FieldNumber, VARIANT * SortFieldType, VARIANT * SortOrder, VARIANT * FieldNumber2, VARIANT * SortFieldType2, VARIANT * SortOrder2, VARIANT * FieldNumber3, VARIANT * SortFieldType3, VARIANT * SortOrder3, VARIANT * SortColumn, VARIANT * Separator, VARIANT * CaseSensitive, VARIANT * LanguageID)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xa8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, LanguageID);
    }
    void SortAscending()
    {
        InvokeHelper(0xa9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SortDescending()
    {
        InvokeHelper(0xaa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    BOOL IsEqual(LPDISPATCH Range)
    {
        BOOL result;
        static BYTE parms[] = VTS_DISPATCH;
        InvokeHelper(0xab, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Range);
        return result;
    }
    float Calculate()
    {
        float result;
        InvokeHelper(0xac, DISPATCH_METHOD, VT_R4, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH GoTo(VARIANT * What, VARIANT * Which, VARIANT * Count, VARIANT * Name)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xad, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, What, Which, Count, Name);
        return result;
    }
    LPDISPATCH GoToNext(long What)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0xae, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, What);
        return result;
    }
    LPDISPATCH GoToPrevious(long What)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0xaf, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, What);
        return result;
    }
    void PasteSpecial(VARIANT * IconIndex, VARIANT * Link, VARIANT * Placement, VARIANT * DisplayAsIcon, VARIANT * DataType, VARIANT * IconFileName, VARIANT * IconLabel)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xb0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel);
    }
    void LookupNameProperties()
    {
        InvokeHelper(0xb1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    long ComputeStatistics(long Statistic)
    {
        long result;
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0xb2, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Statistic);
        return result;
    }
    void Relocate(long Direction)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0xb3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Direction);
    }
    void CheckSynonyms()
    {
        InvokeHelper(0xb4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void SubscribeTo(LPCTSTR Edition, VARIANT * Format)
    {
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT;
        InvokeHelper(0xb5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Edition, Format);
    }
    void CreatePublisher(VARIANT * Edition, VARIANT * ContainsPICT, VARIANT * ContainsRTF, VARIANT * ContainsText)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xb6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Edition, ContainsPICT, ContainsRTF, ContainsText);
    }
    void InsertAutoText()
    {
        InvokeHelper(0xb7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void InsertDatabase(VARIANT * Format, VARIANT * Style, VARIANT * LinkToSource, VARIANT * Connection, VARIANT * SQLStatement, VARIANT * SQLStatement1, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate, VARIANT * DataSource, VARIANT * From, VARIANT * To, VARIANT * IncludeFields)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xc2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Format, Style, LinkToSource, Connection, SQLStatement, SQLStatement1, PasswordDocument, PasswordTemplate, WritePasswordDocument, WritePasswordTemplate, DataSource, From, To, IncludeFields);
    }
    void AutoFormat()
    {
        InvokeHelper(0xc3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void CheckGrammar()
    {
        InvokeHelper(0xcc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void CheckSpelling(VARIANT * CustomDictionary, VARIANT * IgnoreUppercase, VARIANT * AlwaysSuggest, VARIANT * CustomDictionary2, VARIANT * CustomDictionary3, VARIANT * CustomDictionary4, VARIANT * CustomDictionary5, VARIANT * CustomDictionary6, VARIANT * CustomDictionary7, VARIANT * CustomDictionary8, VARIANT * CustomDictionary9, VARIANT * CustomDictionary10)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xcd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, CustomDictionary, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10);
    }
    LPDISPATCH GetSpellingSuggestions(VARIANT * CustomDictionary, VARIANT * IgnoreUppercase, VARIANT * MainDictionary, VARIANT * SuggestionMode, VARIANT * CustomDictionary2, VARIANT * CustomDictionary3, VARIANT * CustomDictionary4, VARIANT * CustomDictionary5, VARIANT * CustomDictionary6, VARIANT * CustomDictionary7, VARIANT * CustomDictionary8, VARIANT * CustomDictionary9, VARIANT * CustomDictionary10)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xd1, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10);
        return result;
    }
    void InsertParagraphBefore()
    {
        InvokeHelper(0xd4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void NextSubdocument()
    {
        InvokeHelper(0xdb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void PreviousSubdocument()
    {
        InvokeHelper(0xdc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ConvertHangulAndHanja(VARIANT * ConversionsMode, VARIANT * FastConversion, VARIANT * CheckHangulEnding, VARIANT * EnableRecentOrdering, VARIANT * CustomDictionary)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xdd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ConversionsMode, FastConversion, CheckHangulEnding, EnableRecentOrdering, CustomDictionary);
    }
    void PasteAsNestedTable()
    {
        InvokeHelper(0xde, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    void ModifyEnclosure(VARIANT * Style, VARIANT * Symbol, VARIANT * EnclosedText)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0xdf, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Style, Symbol, EnclosedText);
    }
    void PhoneticGuide(LPCTSTR Text, long Alignment, long Raise, long FontSize, LPCTSTR FontName)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4 VTS_I4 VTS_BSTR;
        InvokeHelper(0xe0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text, Alignment, Raise, FontSize, FontName);
    }
    void InsertDateTime(VARIANT * DateTimeFormat, VARIANT * InsertAsField, VARIANT * InsertAsFullWidth, VARIANT * DateLanguage, VARIANT * CalendarType)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1bc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType);
    }
    void Sort(VARIANT * ExcludeHeader, VARIANT * FieldNumber, VARIANT * SortFieldType, VARIANT * SortOrder, VARIANT * FieldNumber2, VARIANT * SortFieldType2, VARIANT * SortOrder2, VARIANT * FieldNumber3, VARIANT * SortFieldType3, VARIANT * SortOrder3, VARIANT * SortColumn, VARIANT * Separator, VARIANT * CaseSensitive, VARIANT * BidiSort, VARIANT * IgnoreThe, VARIANT * IgnoreKashida, VARIANT * IgnoreDiacritics, VARIANT * IgnoreHe, VARIANT * LanguageID)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1e4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID);
    }
    void DetectLanguage()
    {
        InvokeHelper(0xcb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH ConvertToTable(VARIANT * Separator, VARIANT * NumRows, VARIANT * NumColumns, VARIANT * InitialColumnWidth, VARIANT * Format, VARIANT * ApplyBorders, VARIANT * ApplyShading, VARIANT * ApplyFont, VARIANT * ApplyColor, VARIANT * ApplyHeadingRows, VARIANT * ApplyLastRow, VARIANT * ApplyFirstColumn, VARIANT * ApplyLastColumn, VARIANT * AutoFit, VARIANT * AutoFitBehavior, VARIANT * DefaultTableBehavior)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1f2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior);
        return result;
    }
    void TCSCConverter(long WdTCSCConverterDirection, BOOL CommonTerms, BOOL UseVariants)
    {
        static BYTE parms[] = VTS_I4 VTS_BOOL VTS_BOOL;
        InvokeHelper(0x1f3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, WdTCSCConverterDirection, CommonTerms, UseVariants);
    }
    BOOL get_LanguageDetected()
    {
        BOOL result;
        InvokeHelper(0x107, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_LanguageDetected(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x107, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    float get_FitTextWidth()
    {
        float result;
        InvokeHelper(0x108, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
        return result;
    }
    void put_FitTextWidth(float newValue)
    {
        static BYTE parms[] = VTS_R4;
        InvokeHelper(0x108, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_HorizontalInVertical()
    {
        long result;
        InvokeHelper(0x109, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_HorizontalInVertical(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x109, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_TwoLinesInOne()
    {
        long result;
        InvokeHelper(0x10a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_TwoLinesInOne(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x10a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    BOOL get_CombineCharacters()
    {
        BOOL result;
        InvokeHelper(0x10b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_CombineCharacters(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x10b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_NoProofing()
    {
        long result;
        InvokeHelper(0x143, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_NoProofing(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x143, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_TopLevelTables()
    {
        LPDISPATCH result;
        InvokeHelper(0x144, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Scripts()
    {
        LPDISPATCH result;
        InvokeHelper(0x145, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_CharacterWidth()
    {
        long result;
        InvokeHelper(0x146, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_CharacterWidth(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x146, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_Kana()
    {
        long result;
        InvokeHelper(0x147, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_Kana(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x147, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_BoldBi()
    {
        long result;
        InvokeHelper(0x190, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_BoldBi(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x190, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    long get_ItalicBi()
    {
        long result;
        InvokeHelper(0x191, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void put_ItalicBi(long newValue)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x191, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    CString get_ID()
    {
        CString result;
        InvokeHelper(0x195, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void put_ID(LPCTSTR newValue)
    {
        static BYTE parms[] = VTS_BSTR;
        InvokeHelper(0x195, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_HTMLDivisions()
    {
        LPDISPATCH result;
        InvokeHelper(0x196, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_SmartTags()
    {
        LPDISPATCH result;
        InvokeHelper(0x197, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    BOOL get_ShowAll()
    {
        BOOL result;
        InvokeHelper(0x198, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
        return result;
    }
    void put_ShowAll(BOOL newValue)
    {
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x198, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
    }
    LPDISPATCH get_Document()
    {
        LPDISPATCH result;
        InvokeHelper(0x199, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_FootnoteOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x19a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_EndnoteOptions()
    {
        LPDISPATCH result;
        InvokeHelper(0x19b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void PasteAndFormat(long Type)
    {
        static BYTE parms[] = VTS_I4;
        InvokeHelper(0x19c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type);
    }
    void PasteExcelTable(BOOL LinkedToExcel, BOOL WordFormatting, BOOL RTF)
    {
        static BYTE parms[] = VTS_BOOL VTS_BOOL VTS_BOOL;
        InvokeHelper(0x19d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, LinkedToExcel, WordFormatting, RTF);
    }
    void PasteAppendTable()
    {
        InvokeHelper(0x19e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
    }
    LPDISPATCH get_XMLNodes()
    {
        LPDISPATCH result;
        InvokeHelper(0x154, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_XMLParentNode()
    {
        LPDISPATCH result;
        InvokeHelper(0x155, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Editors()
    {
        LPDISPATCH result;
        InvokeHelper(0x157, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    CString get_XML(BOOL DataOnly)
    {
        CString result;
        static BYTE parms[] = VTS_BOOL;
        InvokeHelper(0x158, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, DataOnly);
        return result;
    }
    VARIANT get_EnhMetaFileBits()
    {
        VARIANT result;
        InvokeHelper(0x159, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH GoToEditableRange(VARIANT * EditorID)
    {
        LPDISPATCH result;
        static BYTE parms[] = VTS_PVARIANT;
        InvokeHelper(0x19f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, EditorID);
        return result;
    }
    void InsertXML(LPCTSTR XML, VARIANT * Transform)
    {
        static BYTE parms[] = VTS_BSTR VTS_PVARIANT;
        InvokeHelper(0x1a0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, XML, Transform);
    }
    void InsertCaption(VARIANT * Label, VARIANT * Title, VARIANT * TitleAutoText, VARIANT * Position, VARIANT * ExcludeLabel)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1a1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Label, Title, TitleAutoText, Position, ExcludeLabel);
    }
    void InsertCrossReference(VARIANT * ReferenceType, long ReferenceKind, VARIANT * ReferenceItem, VARIANT * InsertAsHyperlink, VARIANT * IncludePosition, VARIANT * SeparateNumbers, VARIANT * SeparatorString)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1a2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString);
    }
    LPDISPATCH get_OMaths()
    {
        LPDISPATCH result;
        InvokeHelper(0x15a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    VARIANT get_CharacterStyle()
    {
        VARIANT result;
        InvokeHelper(0x1a4, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    VARIANT get_ParagraphStyle()
    {
        VARIANT result;
        InvokeHelper(0x1a5, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    VARIANT get_ListStyle()
    {
        VARIANT result;
        InvokeHelper(0x1a6, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    VARIANT get_TableStyle()
    {
        VARIANT result;
        InvokeHelper(0x1a7, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_ContentControls()
    {
        LPDISPATCH result;
        InvokeHelper(0x1a8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ExportFragment(LPCTSTR FileName, long Format)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4;
        InvokeHelper(0x1a9, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, Format);
    }
    CString get_WordOpenXML()
    {
        CString result;
        InvokeHelper(0x1aa, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
        return result;
    }
    void SetListLevel(short Level)
    {
        static BYTE parms[] = VTS_I2;
        InvokeHelper(0x1ab, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Level);
    }
    void InsertAlignmentTab(long Alignment, long RelativeTo)
    {
        static BYTE parms[] = VTS_I4 VTS_I4;
        InvokeHelper(0x1f4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Alignment, RelativeTo);
    }
    LPDISPATCH get_ParentContentControl()
    {
        LPDISPATCH result;
        InvokeHelper(0x1f5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    void ImportFragment(LPCTSTR FileName, BOOL MatchDestination)
    {
        static BYTE parms[] = VTS_BSTR VTS_BOOL;
        InvokeHelper(0x1f6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName, MatchDestination);
    }
    void ExportAsFixedFormat(LPCTSTR OutputFileName, long ExportFormat, BOOL OpenAfterExport, long OptimizeFor, BOOL ExportCurrentPage, long Item, BOOL IncludeDocProps, BOOL KeepIRM, long CreateBookmarks, BOOL DocStructureTags, BOOL BitmapMissingFonts, BOOL UseISO19005_1, VARIANT * FixedFormatExtClassPtr)
    {
        static BYTE parms[] = VTS_BSTR VTS_I4 VTS_BOOL VTS_I4 VTS_BOOL VTS_I4 VTS_BOOL VTS_BOOL VTS_I4 VTS_BOOL VTS_BOOL VTS_BOOL VTS_PVARIANT;
        InvokeHelper(0x1f7, DISPATCH_METHOD, VT_EMPTY, NULL, parms, OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr);
    }
    LPDISPATCH get_Locks()
    {
        LPDISPATCH result;
        InvokeHelper(0x1f8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Updates()
    {
        LPDISPATCH result;
        InvokeHelper(0x1f9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    LPDISPATCH get_Conflicts()
    {
        LPDISPATCH result;
        InvokeHelper(0x1fa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
        return result;
    }
    long get_TextVisibleOnScreen()
    {
        long result;
        InvokeHelper(0x1fb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
        return result;
    }
    void SortByHeadings(VARIANT * SortFieldType, VARIANT * SortOrder, VARIANT * CaseSensitive, VARIANT * BidiSort, VARIANT * IgnoreThe, VARIANT * IgnoreKashida, VARIANT * IgnoreDiacritics, VARIANT * IgnoreHe, VARIANT * LanguageID)
    {
        static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
        InvokeHelper(0x1fc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SortFieldType, SortOrder, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID);
    }

    // Range 属性
public:

};

}

#endif // _WORD_INTERFACES_H_