
// OfficeConverterDlg.h : 头文件
//

#pragma once


// COfficeConverterDlg 对话框
class COfficeConverterDlg : public CDialogEx
{
// 构造
public:
	COfficeConverterDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_OFFICECONVERTER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

    afx_msg void OnBnClickedButtonPath();
public:
    afx_msg void OnBnClickedButton1();
};
