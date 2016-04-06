
// OfficeConverterDlg.cpp : ʵ���ļ�
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


// COfficeConverterDlg �Ի���



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


// COfficeConverterDlg ��Ϣ�������

BOOL COfficeConverterDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO:  �ڴ���Ӷ���ĳ�ʼ������

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void COfficeConverterDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR COfficeConverterDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void COfficeConverterDlg::OnBnClickedButtonPath()
{
    // TODO: �ڴ���ӿؼ�֪ͨ����������
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
        MessageBox(L"��ѡ���ļ���");
        return;
    }
    if (!PathFileExists(strFilePath))
    {
        MessageBox(L"�ļ������ڣ�");
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
        AfxMessageBox(L"�޷�ת�������ͣ�");
        return;
    }
    bool result = converter->Convert(strFilePath.GetBuffer(),
                                     szOutDir.GetBuffer());
    converter.reset();
    strFilePath.ReleaseBuffer();
    szOutDir.ReleaseBuffer();
    if (result)
    {
        AfxMessageBox(L"ת����ɣ�ͼƬ�ļ�������Դ�ļ�ͬĿ¼�¡�");
    }
    else
    {
        AfxMessageBox(L"ת��ʧ�ܣ�");
    }
}
