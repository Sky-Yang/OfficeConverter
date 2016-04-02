
// OfficeConverterDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "OfficeConverter.h"

#include <memory>

#include "OfficeConverterDlg.h"
#include "afxdialogex.h"

#include "office/office_converter.h"
#include "office/word/word_converter.h"
#include "office/excel/excel_converter.h"
#include "office/ppt/ppt_converter.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// COfficeConverterDlg 对话框



COfficeConverterDlg::COfficeConverterDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(COfficeConverterDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void COfficeConverterDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(COfficeConverterDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
    ON_BN_CLICKED(IDC_BUTTON_PATH, &COfficeConverterDlg::OnBnClickedButtonPath)
    ON_BN_CLICKED(IDC_BUTTON1, &COfficeConverterDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// COfficeConverterDlg 消息处理程序

BOOL COfficeConverterDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO:  在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void COfficeConverterDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR COfficeConverterDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void COfficeConverterDlg::OnBnClickedButtonPath()
{
    // TODO: 在此添加控件通知处理程序代码
    CFileDialog dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY,
                        L"office file(*.doc,*.ppt,*.xls,*.docx,*.pptx,*.xlsx,*.pdf)|*.doc;*.ppt;*.xls;*.pdf;*.docx;*.pptx;*.xlsx;||", 
                        CWnd::FromHandle(m_hWnd));

    if (dlgFile.DoModal() == IDOK)
    {
        CString	strFilePath = dlgFile.GetPathName();
        CString szFileExt = dlgFile.GetFileExt();
        CString strFileName = dlgFile.GetFileName();

        GetDlgItem(IDC_EDIT_PATH)->SetWindowText(strFilePath);
    }
}


void COfficeConverterDlg::OnBnClickedButton1()
{
    CString	strFilePath;
    GetDlgItem(IDC_EDIT_PATH)->GetWindowText(strFilePath);
    if (strFilePath.IsEmpty())
    {
        MessageBox(L"请选择文件！");
        return;
    }
    if (!PathFileExists(strFilePath))
    {
        MessageBox(L"文件不存在！");
        return;
    }

    CString szFileExt = strFilePath.Right(
        strFilePath.GetLength() - strFilePath.ReverseFind(L'.') - 1);
    szFileExt.MakeLower();

    CString szOutDir;
    szOutDir.Format(L"%s",
                    strFilePath.Left(strFilePath.Find(L".")));

    std::shared_ptr<OfficeConverter> converter;
    if (szFileExt == L"doc" || szFileExt == L"docx")
    {
        converter.reset(new WordConverter());
    }
    else if (szFileExt == L"xls" || szFileExt == L"xlsx")
    {
        converter.reset(new ExcelConverter());
    }
    else if (szFileExt == L"ppt" || szFileExt == L"pptx")
    {
        converter.reset(new PptConverter(1280, 720));
    }
    else
    {
        AfxMessageBox(L"无法转换的类型！");
        return;
    }
    bool result = converter->Convert(strFilePath.GetBuffer(),
                                     szOutDir.GetBuffer());
    converter.reset();
    strFilePath.ReleaseBuffer();
    szOutDir.ReleaseBuffer();
    if (result)
    {
        AfxMessageBox(L"转换完成，图片文件保存在源文件同目录下。");
    }
    else
    {
        AfxMessageBox(L"转换失败！");
    }
}
