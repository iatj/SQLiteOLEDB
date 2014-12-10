
#pragma warning( disable: 4049 )  /* more than 64k source lines */

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 6.00.0347 */
/* at Wed Oct 16 14:42:38 2002
 */
/* Compiler settings for C:\DOCUME~1\ADMINI~1\LOCALS~1\Temp\IDL16.tmp:
    Os, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#if !defined(_M_IA64) && !defined(_M_AMD64)

#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, LIBID_VBScript_RegExp_55,0x3F4DACA7,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IRegExp,0x3F4DACA0,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IMatch,0x3F4DACA1,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IMatchCollection,0x3F4DACA2,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IRegExp2,0x3F4DACB0,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IMatch2,0x3F4DACB1,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IMatchCollection2,0x3F4DACB2,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_ISubMatches,0x3F4DACB3,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(CLSID, CLSID_RegExp,0x3F4DACA4,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(CLSID, CLSID_Match,0x3F4DACA5,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(CLSID, CLSID_MatchCollection,0x3F4DACA6,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(CLSID, CLSID_SubMatches,0x3F4DACC0,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



#endif /* !defined(_M_IA64) && !defined(_M_AMD64)*/


#pragma warning( disable: 4049 )  /* more than 64k source lines */

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 6.00.0347 */
/* at Wed Oct 16 14:42:38 2002
 */
/* Compiler settings for C:\DOCUME~1\ADMINI~1\LOCALS~1\Temp\IDL16.tmp:
    Oicf, W1, Zp8, env=Win64 (32b run,appending)
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#if defined(_M_IA64) || defined(_M_AMD64)

#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, LIBID_VBScript_RegExp_55,0x3F4DACA7,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IRegExp,0x3F4DACA0,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IMatch,0x3F4DACA1,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IMatchCollection,0x3F4DACA2,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IRegExp2,0x3F4DACB0,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IMatch2,0x3F4DACB1,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_IMatchCollection2,0x3F4DACB2,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(IID, IID_ISubMatches,0x3F4DACB3,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(CLSID, CLSID_RegExp,0x3F4DACA4,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(CLSID, CLSID_Match,0x3F4DACA5,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(CLSID, CLSID_MatchCollection,0x3F4DACA6,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);


MIDL_DEFINE_GUID(CLSID, CLSID_SubMatches,0x3F4DACC0,0x160D,0x11D2,0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



#endif /* defined(_M_IA64) || defined(_M_AMD64)*/

