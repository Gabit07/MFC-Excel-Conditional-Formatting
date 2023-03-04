
// DlgProxy.h: header file
//

#pragma once

class CMFCApplication1Dlg;


// CMFCApplication1DlgAutoProxy command target

class CMFCApplication1DlgAutoProxy : public CCmdTarget
{
	DECLARE_DYNCREATE(CMFCApplication1DlgAutoProxy)

	CMFCApplication1DlgAutoProxy();           // protected constructor used by dynamic creation

// Attributes
public:
	CMFCApplication1Dlg* m_pDialog;

// Operations
public:

// Overrides
	public:
	virtual void OnFinalRelease();

// Implementation
protected:
	virtual ~CMFCApplication1DlgAutoProxy();

	// Generated message map functions

	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CMFCApplication1DlgAutoProxy)

	// Generated OLE dispatch map functions

	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

