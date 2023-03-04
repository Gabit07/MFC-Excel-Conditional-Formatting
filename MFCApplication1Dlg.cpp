
// MFCApplication1Dlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "MFCApplication1.h"
#include "MFCApplication1Dlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CMFCApplication1Dlg dialog


IMPLEMENT_DYNAMIC(CMFCApplication1Dlg, CDialogEx);

CMFCApplication1Dlg::CMFCApplication1Dlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_MFCAPPLICATION1_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = nullptr;
}

CMFCApplication1Dlg::~CMFCApplication1Dlg()
{
	// If there is an automation proxy for this dialog, set
	//  its back pointer to this dialog to null, so it knows
	//  the dialog has been deleted.
	if (m_pAutoProxy != nullptr)
		m_pAutoProxy->m_pDialog = nullptr;
}

void CMFCApplication1Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CMFCApplication1Dlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CMFCApplication1Dlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CMFCApplication1Dlg message handlers

BOOL CMFCApplication1Dlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CMFCApplication1Dlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMFCApplication1Dlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CMFCApplication1Dlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// Automation servers should not exit when a user closes the UI
//  if a controller still holds on to one of its objects.  These
//  message handlers make sure that if the proxy is still in use,
//  then the UI is hidden but the dialog remains around if it
//  is dismissed.

void CMFCApplication1Dlg::OnClose()
{
	if (CanExit())
		CDialogEx::OnClose();
}

void CMFCApplication1Dlg::OnOK()
{
	if (CanExit())
		CDialogEx::OnOK();
}

void CMFCApplication1Dlg::OnCancel()
{
	if (CanExit())
		CDialogEx::OnCancel();
}

BOOL CMFCApplication1Dlg::CanExit()
{
	// If the proxy object is still around, then the automation
	//  controller is still holding on to this application.  Leave
	//  the dialog around, but hide its UI.
	if (m_pAutoProxy != nullptr)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}


void CMFCApplication1Dlg::OnBnClickedButton1()
{
	// TODO: Add your control notification handler code here
	CRange range(nullptr);
	CApplication app(nullptr);
	CWorkbook wbook(nullptr);
	CWorkbooks wbooks(nullptr);
	CWorksheet wsheet(nullptr);
	CWorksheets		wsheets(nullptr);
	COleVariant covTrue((short)TRUE);
	COleVariant    covFalse((short)FALSE);
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CString strText;

	if (!app.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("Could not start Excel."));
	}

	// workbook 생성
	app.put_SheetsInNewWorkbook(1);

	// books 연결
	wbooks.AttachDispatch(app.get_Workbooks());

	// book 연결
	wbook.AttachDispatch(wbooks.Add(covOptional));

	// sheets 생성, 연결
	wsheets = wbook.get_Sheets();

	// sheet 생성, 연결 (1번 시트)
	wsheet = wsheets.get_Item(COleVariant((short)1));

	// sheet 이름 변경
	wsheet.put_Name(_T("Cam1"));

	// range 생성, 연결
	range.AttachDispatch(wsheet.get_Cells(), true);

	// Insert Data
	for (int y = 0; y <= 40; y++) // => y
	{
		for (int x = 0; x <= 40; x++) // => x
		{
			int nIQvalue = 110 + x;
			strText.Format(_T("%d"), nIQvalue);
			range.put_Item(COleVariant((long)(y + 1)), COleVariant((long)(x + 1)), COleVariant(strText));
		}
	}

	//눈금선 없애기
	CWindow0 window = app.get_ActiveWindow();
	window.put_DisplayGridlines(false);

	// Font
	CFont0 font = range.get_Font();
	font.put_Size(COleVariant(12L));

	range.put_RowHeight(COleVariant((double)15));
	range.put_ColumnWidth(COleVariant((double)4));

	CFormatConditions formatConditions = range.get_FormatConditions();
	CColorScale colorscale = formatConditions.AddColorScale(3);

	CColorScaleCriteria colsc = colorscale.get_ColorScaleCriteria();

	CColorScaleCriterion csn = colsc.get_Item(COleVariant((short)1));
	csn.put_Type((long)0);
	csn.put_Value(COleVariant((short)110));

	CFormatColor formatcolor1 = csn.get_FormatColor();
	formatcolor1.put_Color(COleVariant((double)RGB(255, 235, 132)));

	CColorScaleCriterion csn2 = colsc.get_Item(COleVariant((short)2));
	csn2.put_Type((long)0);
	csn2.put_Value(COleVariant((short)130));
	CFormatColor formatcolor2 = csn2.get_FormatColor();
	formatcolor2.put_Color(COleVariant((double)RGB(99, 190, 123)));

	CColorScaleCriterion csn3 = colsc.get_Item(COleVariant((short)3));
	csn3.put_Type((long)0);
	csn3.put_Value(COleVariant((short)150));
	CFormatColor formatcolor3 = csn3.get_FormatColor();
	formatcolor3.put_Color(COleVariant((double)RGB(248, 105, 107)));

	window.put_Zoom(COleVariant((short)85));

	//열리는 과정이 보이게
	app.put_Visible(TRUE);
	app.put_DisplayAlerts(VARIANT_FALSE);
	app.put_UserControl(TRUE);

	// 연결 끊기
	range.ReleaseDispatch();
	wsheet.ReleaseDispatch();
	wsheets.ReleaseDispatch();
	wbook.ReleaseDispatch();
	wbooks.ReleaseDispatch();
	app.ReleaseDispatch();
}
