// 컴퓨터에서 형식 라이브러리 마법사의 [클래스 추가]를 사용하여 생성한 IDispatch 래퍼 클래스입니다.

//#import "C:\\Program Files\\Microsoft Office\\Office14\\EXCEL.EXE" no_namespace
// CColorScaleCriteria 래퍼 클래스

class CColorScaleCriteria : public COleDispatchDriver
{
public:
	CColorScaleCriteria() {} // COleDispatchDriver 기본 생성자를 호출합니다.
	CColorScaleCriteria(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CColorScaleCriteria(const CColorScaleCriteria& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 특성
public:

	// 작업
public:


	// ColorScaleCriteria 메서드
public:
	long get_Count()
	{
		long result;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get__Default(VARIANT& Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Item(VARIANT& Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}

	// ColorScaleCriteria 속성
public:

};
