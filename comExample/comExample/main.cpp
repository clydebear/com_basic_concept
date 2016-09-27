#include <windows.h>
#include <objbase.h>
/*com������һ�ֹ淶������ʵ�֣���ʹ��c++ʵ��ʱ��com�������һ��c++�࣬���ӿ��Ǵ��麯��,
���´��������һ��com���
class Ifunction   //������ǽӿ�
{
public:
virtual Method1(...) = 0;
virtual Method2(...) = 0;
//.....
}

class MyObject:public IFunction   //�������com���
{
public:
virtual Method1(...){...}
virtual Method(,,,){,,,}
//....
}

com�淶�涨�κ������ӿڶ���Ҫ��IUnknown�ӿ��м̳ж�����IUnknown�����һ��3����Ҫ������
QueryInterface ----������������ϵĽӿڲ�ѯ
AddRef ----�����������ü���
Release----���ڼ������ô��� �����ü���������com�������������

����һ����Ҫ�Ľӿ� IClassFactory�������ⲿʹ������˵com�������һ�㲻��֪���ʹ涨ÿ�����������ʵ��һ����֮���Ӧ��
�๤����class factoryҲ��һ��com���������IClassFactory�Ľӿں���CreateInstance �У�����ʹ��new��������һ��com��������ʵ��

com���һ�����������ͣ���������������ؽ��������Զ�����
ÿ��Com�����ʹ��һ��GUID����CLSID_Object����Ψһ��ʶ��
////////////////////////////////////////////////////////////////////////////////
����com������̣���ͨ��GUID����CoGetClassObject����ô����������Ķ�����๤����Ȼ������๤���Ľӿڷ���IClassFactory::CreateInstance�ʹ�����һ��CLSID_Object��ʶ�����������

CoGetClassObject�Ĺ��̣�
CoGetClassObject����������
{
  1ͨ����ѯע���CLSID_Object��֪���DLL�ļ�·��
  2װ��DLL�⣨����LoadLibrary��
  3ʹ�ú���GetProcAddress�����������õ�DLL�к���DLLGetClassObJect�ĺ���ָ��
  4����DllGetClassObject�õ��๤������ָ��
}

DllGetClassObject�Ĺ��̣�
DllGetClassObject��...��  //�����Ǹ���ָ�������GUID������Ӧ���๤�����󣬲���������๤����IClassFactory�ӿ�
{
     //�����๤������
	CFactory *pFactory = Factory;
	//��ѯ�õ�IClassFactoryָ��
	pFactory->QueryInterface(IID_IClassFactory,(void**)&pClassFactory);
	pFactory->release();
	//...
}

CFactory::CreateInstance(..)    //��IClassFactory�ӿڵ�һ���ӿڷ������������մ����������ʵ����
{
//..
//����CLSID_Object��Ӧ���������
CObject* pObject = new COject;   //��������������ڴ�
pObject->QueryInterface(IID_IUNkown,(void**)&pUnk);
pObject->Release();
//....
}

*/

int WINAPI WinMain(HINSTANCE hInstance,HINSTANCE hPrevInstance,LPSTR lpCmdLine,int nShowCmd)

{
	//com���ʼ��
	::CoInitialize(NULL);
	IUnknown *pUnk = NULL;
	IObject *pObject = NULL;
	//��������
	HRESULT hr = CoCreateInstance(CLSID_Object,CLSCTX_INPROC_SERVER,NULL,IID_IUnknown,(void**)&pUnk);

	if (SUCCEEDED(hr))
	{
		hr = pUnk->QueryInterface(IID_IObject,(void **)&pObject);
		if (SUCCEEDED(hr))
		{
			//���ýӿڷ���
			pObject->SomeMethod();
			pObject->Release();
		}
		pUnk->Release(); //�ͷŽӿ�
	}

	CoUninitialize();//�ͷ�com��ʹ�õ���Դ
	return 0;
}