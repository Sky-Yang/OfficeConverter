
// OfficeConverterDlg.h : ͷ�ļ�
//

#pragma once


// COfficeConverterDlg �Ի���
class COfficeConverterDlg : public CDialogEx
{
// ����
public:
	COfficeConverterDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_OFFICECONVERTER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

    afx_msg void OnBnClickedButtonPath();
public:
    afx_msg void OnBnClickedButton1();
};
