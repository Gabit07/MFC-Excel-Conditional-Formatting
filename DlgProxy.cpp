
// DlgProxy.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "MFCApplication1.h"
#include "DlgProxy.h"
#include "MFCApplication1Dlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CMFCApplication1DlgAutoProxy

IMPLEMENT_DYNCREATE(CMFCApplication1DlgAutoProxy, CCmdTarget)

CMFCApplication1DlgAutoProxy::CMFCApplication1DlgAutoProxy()
{
	EnableAutomation();

	// To keep the application running as long as an automation
	//	object is active, the constructor calls AfxOleLockApp.
	AfxOleLockApp();

	// Get access to the dialog through the application's
	//  main window pointer.  Set the proxy's internal pointer
	//  to point to the dialog, and set the dialog's back pointer to
	//  this proxy.
	ASSERT_VALID(AfxGetApp()->m_pMainWnd);
	if (AfxGetApp()->m_pMainWnd)
	{
		ASSERT_KINDOF(CMFCApplication1Dlg, AfxGetApp()->m_pMainWnd);
		if (AfxGetApp()->m_pMainWnd->IsKindOf(RUNTIME_CLASS(CMFCApplication1Dlg)))
		{
			m_pDialog = reinterpret_cast<CMFCApplication1Dlg*>(AfxGetApp()->m_pMainWnd);
			m_pDialog->m_pAutoProxy = this;
		}
	}
}

CMFCApplication1DlgAutoProxy::~CMFCApplication1DlgAutoProxy()
{
	// To terminate the application when all objects created with
	// 	with automation, the destructor calls AfxOleUnlockApp.
	//  Among other things, this will destroy the main dialog
	if (m_pDialog != nullptr)
		m_pDialog->m_pAutoProxy = nullptr;
	AfxOleUnlockApp();
}

void CMFCApplication1DlgAutoProxy::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CMFCApplication1DlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CMFCApplication1DlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// Note: we add support for IID_IMFCApplication1 to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the
//  dispinterface in the .IDL file.

// {2630887f-970a-4f36-9082-12b15d1c0f42}
static const IID IID_IMFCApplication1 =
{0x2630887f,0x970a,0x4f36,{0x90,0x82,0x12,0xb1,0x5d,0x1c,0x0f,0x42}};

BEGIN_INTERFACE_MAP(CMFCApplication1DlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CMFCApplication1DlgAutoProxy, IID_IMFCApplication1, Dispatch)
END_INTERFACE_MAP()

// The IMPLEMENT_OLECREATE2 macro is defined in pch.h of this project
// {dd223745-877f-4c92-b5a7-19c5b2b59613}
IMPLEMENT_OLECREATE2(CMFCApplication1DlgAutoProxy, "MFCApplication1.Application", 0xdd223745,0x877f,0x4c92,0xb5,0xa7,0x19,0xc5,0xb2,0xb5,0x96,0x13)


// CMFCApplication1DlgAutoProxy message handlers
