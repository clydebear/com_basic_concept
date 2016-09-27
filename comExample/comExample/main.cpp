#include <windows.h>
#include <objbase.h>
/*com本身是一种规范而不是实现，当使用c++实现时，com组件就是一个c++类，而接口是纯虚函数,
如下代码简单描述一个com组件
class Ifunction   //这个就是接口
{
public:
virtual Method1(...) = 0;
virtual Method2(...) = 0;
//.....
}

class MyObject:public IFunction   //这个就是com组件
{
public:
virtual Method1(...){...}
virtual Method(,,,){,,,}
//....
}

com规范规定任何组件或接口都需要从IUnknown接口中继承而来。IUnknown组件定一个3个重要函数：
QueryInterface ----负责组件对象上的接口查询
AddRef ----用于增加引用计数
Release----用于减少引用次数 ，引用计数决定了com对象的生命周期

另外一个重要的接口 IClassFactory：对于外部使用者来说com组件类名一般不可知，故规定每个组件都必须实现一个与之相对应的
类工厂（class factory也是一个com组件），在IClassFactory的接口函数CreateInstance 中，才能使用new操作生成一个com组件类对象实例

com组件一般有三种类型：进程内组件、本地进程组件和远程组件
每个Com组件都使用一个GUID（即CLSID_Object）来唯一标识。
////////////////////////////////////////////////////////////////////////////////
创建com组件过程：先通过GUID调用CoGetClassObject来获得创建这个组件的对象的类工厂，然后调用类工厂的接口方法IClassFactory::CreateInstance就创建了一个CLSID_Object标识的组件对象了

CoGetClassObject的过程：
CoGetClassObject（。。。）
{
  1通过查询注册表CLSID_Object得知组件DLL文件路径
  2装入DLL库（调用LoadLibrary）
  3使用函数GetProcAddress（。。。）得到DLL中函数DLLGetClassObJect的函数指针
  4调用DllGetClassObject得到类工厂对象指针
}

DllGetClassObject的过程：
DllGetClassObject（...）  //作用是根据指定的组件GUID创建响应的类工厂对象，并返回这个类工厂的IClassFactory接口
{
     //创建类工厂对象
	CFactory *pFactory = Factory;
	//查询得到IClassFactory指针
	pFactory->QueryInterface(IID_IClassFactory,(void**)&pClassFactory);
	pFactory->release();
	//...
}

CFactory::CreateInstance(..)    //是IClassFactory接口的一个接口方法，负责最终创建组件对象实例，
{
//..
//创建CLSID_Object对应的组件对象
CObject* pObject = new COject;   //真正的组件功能在此
pObject->QueryInterface(IID_IUNkown,(void**)&pUnk);
pObject->Release();
//....
}

*/

int WINAPI WinMain(HINSTANCE hInstance,HINSTANCE hPrevInstance,LPSTR lpCmdLine,int nShowCmd)

{
	//com库初始化
	::CoInitialize(NULL);
	IUnknown *pUnk = NULL;
	IObject *pObject = NULL;
	//创建对象
	HRESULT hr = CoCreateInstance(CLSID_Object,CLSCTX_INPROC_SERVER,NULL,IID_IUnknown,(void**)&pUnk);

	if (SUCCEEDED(hr))
	{
		hr = pUnk->QueryInterface(IID_IObject,(void **)&pObject);
		if (SUCCEEDED(hr))
		{
			//调用接口方法
			pObject->SomeMethod();
			pObject->Release();
		}
		pUnk->Release(); //释放接口
	}

	CoUninitialize();//释放com库使用的资源
	return 0;
}