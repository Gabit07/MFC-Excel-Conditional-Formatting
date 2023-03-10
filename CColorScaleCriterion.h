// 컴퓨터에서 형식 라이브러리 마법사의 [클래스 추가]를 사용하여 생성한 IDispatch 래퍼 클래스입니다.

//#import "C:\\Program Files\\Microsoft Office\\Office14\\EXCEL.EXE" no_namespace
// CColorScaleCriterion 래퍼 클래스

class CColorScaleCriterion : public COleDispatchDriver
{
public:
	CColorScaleCriterion() {} // COleDispatchDriver 기본 생성자를 호출합니다.
	CColorScaleCriterion(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CColorScaleCriterion(const CColorScaleCriterion& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 특성
public:

	// 작업
public:


	// ColorScaleCriterion 메서드
public:
	long get_Index()
	{
		long result;
		InvokeHelper(0x1e6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Type(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x6c, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	VARIANT get_Value()
	{
		VARIANT result;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, nullptr);
		return result;
	}
	void put_Value(VARIANT& newValue)
	{
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x6, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, &newValue);
	}
	LPDISPATCH get_FormatColor()
	{
		LPDISPATCH result;
		InvokeHelper(0xa9d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}

	// ColorScaleCriterion 속성
public:

};
