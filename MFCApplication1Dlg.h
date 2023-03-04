
// MFCApplication1Dlg.h : header file
//
#include "CApplication.h"
#include "CBorder.h"
#include "CColorScale.h"
#include "CColorScaleCriteria.h"
#include "CColorScaleCriterion.h"
#include "CFont0.h"
#include "CFormatColor.h"
#include "CFormatConditions.h"
#include "CRange.h"
#include "CWindow0.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"


#pragma once

class CMFCApplication1DlgAutoProxy;


// CMFCApplication1Dlg dialog
class CMFCApplication1Dlg : public CDialogEx
{
	DECLARE_DYNAMIC(CMFCApplication1Dlg);
	friend class CMFCApplication1DlgAutoProxy;

// Construction
public:
	CMFCApplication1Dlg(CWnd* pParent = nullptr);	// standard constructor
	virtual ~CMFCApplication1Dlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_MFCAPPLICATION1_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	CMFCApplication1DlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
};
