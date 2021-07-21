Attribute VB_Name = "modDeclares"
Option Explicit
'this is to (start to) eliminate the need for the vblibcurl.tlb
'typelibs put imports right into our import table, so dependancy dlls
'must be found on startup by windows loader, we can't hunt for them
'during dev this can be annoying unless you put everything in system path
'(and remember to update them if they change in dev)


'these enum constants are circa curl 7.13 compiled 2.15.2005
'whats new? current version 15yrs later is 7.73 compiled 10.14.2020




'basic api
'---------------------------------------------------------------------
'[entry(0x60000000), helpstring("Cleanup an easy session")]
'void _stdcall vbcurl_easy_cleanup([in] long easy);
Public Declare Sub vbcurl_easy_cleanup Lib "vblibcurl.dll" (ByVal easy As Long)

'[entry(0x60000001), helpstring("Duplicate an easy handle")]
'long _stdcall vbcurl_easy_duphandle([in] long easy);

'[entry(0x60000002), helpstring("Get information on an easy session")]
'CURLcode _stdcall vbcurl_easy_getinfo(
'                [in] long easy,
'                [in] CURLINFO info,
'                [in] VARIANT* pv);
Public Declare Function vbcurl_easy_getinfo Lib "vblibcurl.dll" ( _
    ByVal easy As Long, _
    ByVal info As CURLINFO, _
    ByRef Value As Variant _
) As CURLcode

'[entry(0x60000003), helpstring("Initialize an easy session")]
'long _stdcall vbcurl_easy_init();
Public Declare Function vbcurl_easy_init Lib "vblibcurl.dll" () As Long
'
'If you did not already call curl_global_init, curl_easy_init does it automatically.
'This may be lethal in multi-threaded cases, since curl_global_init is not thread-safe,
'and it may result in resource problems because there is no corresponding cleanup.
'
'You are strongly advised to not allow this automatic behaviour,
'https://curl.se/libcurl/c/curl_easy_init.html
'
'Note: curl_global_init not exposed by vblibcurl and it only calls curl_easy_init



'[entry(0x60000004), helpstring("Perform an easy transfer")]
'CURLcode _stdcall vbcurl_easy_perform([in] long easy);
Public Declare Function vbcurl_easy_perform Lib "vblibcurl.dll" (ByVal easy As Long) As CURLcode


'[entry(0x60000005), helpstring("Reset an easy handle")]
'void _stdcall vbcurl_easy_reset([in] long easy);
Public Declare Sub vbcurl_easy_reset Lib "vblibcurl.dll" (ByVal easy As Long)

'[entry(0x60000006), helpstring("Set option for easy transfer")]
'CURLcode _stdcall vbcurl_easy_setopt(
'                [in] long easy,
'                [in] CURLoption opt,
'                [in] VARIANT* value);
Public Declare Function vbcurl_easy_setopt Lib "vblibcurl.dll" ( _
    ByVal easy As Long, _
    ByVal opt As CURLoption, _
    ByRef Value As Variant _
) As CURLcode


'we have an internal version in vb see curlCode2Text below
'[entry(0x60000007), helpstring("Get a string description of an error code")]
'BSTR _stdcall vbcurl_easy_strerror([in] CURLcode err);



'forms (complete)
'---------------------------------------------------------------------
'[entry(0x60000008), helpstring("Add two option/value pairs to a form part")]
'CURLFORMcode _stdcall vbcurl_form_add_four_to_part(
'                [in] long part,
'                [in] CURLformoption opt1,
'                [in] VARIANT* val1,
'                [in] CURLformoption opt2,
'                [in] VARIANT* val2);
Public Declare Function vbcurl_form_add_four_to_part Lib "vblibcurl.dll" ( _
    ByVal hPart As Long, _
    ByVal opt1 As CURLformoption, _
    ByRef Name As Variant, _
    ByVal opt2 As CURLformoption, _
    ByRef Value As Variant _
) As CURLFORMcode

'[entry(0x60000009), helpstring("Add an option/value pair to a form part")]
'CURLFORMcode _stdcall vbcurl_form_add_pair_to_part(
'                [in] long part,
'                [in] CURLformoption opt,
'                [in] VARIANT* val);
Public Declare Function vbcurl_form_add_pair_to_part Lib "vblibcurl.dll" ( _
    ByVal hPart As Long, _
    ByVal opt1 As CURLformoption, _
    ByRef field1 As Variant _
) As CURLFORMcode

'[entry(0x6000000a), helpstring("Add a completed part to a multi-part form")]
'CURLFORMcode _stdcall vbcurl_form_add_part(
'                [in] long form,
'                [in] long part);
Public Declare Function vbcurl_form_add_part Lib "vblibcurl.dll" ( _
    ByVal hForm As Long, _
    ByVal hPart As Long _
) As CURLFORMcode



'[entry(0x6000000b), helpstring("Add three option/value pairs to a form part")]
'CURLFORMcode _stdcall vbcurl_form_add_six_to_part(
'                [in] long part,
'                [in] CURLformoption opt1,
'                [in] VARIANT* val1,
'                [in] CURLformoption opt2,
'                [in] VARIANT* val2,
'                [in] CURLformoption opt3,
'                [in] VARIANT* val3);
Public Declare Function vbcurl_form_add_six_to_part Lib "vblibcurl.dll" ( _
    ByVal hPart As Long, _
    ByVal opt1 As CURLformoption, _
    ByRef field1 As Variant, _
    ByVal opt2 As CURLformoption, _
    ByRef field2 As Variant, _
    ByVal opt3 As CURLformoption, _
    ByRef field3 As Variant _
) As CURLFORMcode


'[entry(0x6000000c), helpstring("Create a multi-part form handle")]
'long _stdcall vbcurl_form_create();
Public Declare Function vbcurl_form_create Lib "vblibcurl.dll" () As Long


'[entry(0x6000000d), helpstring("Create a multi-part form-part")]
'long _stdcall vbcurl_form_create_part([in] long form);
Public Declare Function vbcurl_form_create_part Lib "vblibcurl.dll" (ByVal hForm As Long) As Long


'[entry(0x6000000e), helpstring("Free a form and all its parts")]
'void _stdcall vbcurl_form_free([in] long form);
Public Declare Sub vbcurl_form_free Lib "vblibcurl.dll" (ByVal hForm As Long)




'multi handles (multi threaded downloads)
'---------------------------------------------------------------------
'[entry(0x6000000f), helpstring("Add an easy handle")]
'CURLMcode _stdcall vbcurl_multi_add_handle(
'                [in] long multi,
'                [in] long easy);

'[entry(0x60000010), helpstring("Cleanup a multi handle")]
'CURLMcode _stdcall vbcurl_multi_cleanup([in] long multi);

'[entry(0x60000011), helpstring("Call fdset on internal sockets")]
'CURLMcode _stdcall vbcurl_multi_fdset([in] long multi);

'[entry(0x60000012), helpstring("Read per-easy info for a multi handle")]
'long _stdcall vbcurl_multi_info_read(
'                [in] long multi,
'                [in, out] CURLMSG* msg,
'                [in, out] long* easy,
'                [in, out] CURLcode* code);

'[entry(0x60000013), helpstring("Initialize a multi handle")]
'long _stdcall vbcurl_multi_init();

'[entry(0x60000014), helpstring("Read/write the easy handles")]
'CURLMcode _stdcall vbcurl_multi_perform(
'                [in] long multi,
'                [in, out] long* runningHandles);

'[entry(0x60000015), helpstring("Remove an easy handle")]
'CURLMcode _stdcall vbcurl_multi_remove_handle(
'                [in] long multi,
'                [in] long easy);

'[entry(0x60000016), helpstring("Perform select on easy handles")]
'long _stdcall vbcurl_multi_select(
'                [in] long multi,
'                [in] long timeoutMillis);

'[entry(0x60000017), helpstring("Get a string description of an error code")]
'BSTR _stdcall vbcurl_multi_strerror([in] CURLMcode err);





'string lists (complete) - https://curl.se/libcurl/c/CURLOPT_HTTPHEADER.html
'other places used: Smtp recipients, walking CURLINFO_CERTINFO return info
'---------------------------------------------------------------------
'[entry(0x60000018), helpstring("Append a string to an slist")]
'void _stdcall vbcurl_slist_append(
'                [in] long slist,
'                [in] BSTR str);
Public Declare Sub vbcurl_slist_append Lib "vblibcurl.dll" (ByVal hList As Long, ByVal strPtr As Long)


'[entry(0x60000019), helpstring("Create a string list")]
'long _stdcall vbcurl_slist_create();
Public Declare Function vbcurl_slist_create Lib "vblibcurl.dll" () As Long


'[entry(0x6000001a), helpstring("Free a created string list")]
'void _stdcall vbcurl_slist_free([in] long slist);
Public Declare Sub vbcurl_slist_free Lib "vblibcurl.dll" (ByVal hList As Long)




'(complete)
'we cant return As String because runtime will give us another BSTR with double %00
'so we will have to steal a reference to the one the dll gives us directly
'this is probably a difference between the C api declares mechanism and tlb import mechanism
'---------------------------------------------------------------------
'[entry(0x6000001b), helpstring("Escape an URL")]
'BSTR _stdcall vbcurl_string_escape(
'                [in] BSTR url,
'                [in] long len);
Public Declare Function vbcurl_string_escape Lib "vblibcurl.dll" ( _
    ByVal strPtr As Long, _
    ByVal sz As Long _
) As Long


'[entry(0x6000001c), helpstring("Unescape an URL")]
'BSTR _stdcall vbcurl_string_unescape(
'                [in] BSTR url,
'                [in] long len);
Public Declare Function vbcurl_string_unescape Lib "vblibcurl.dll" ( _
    ByVal strPtr As Long, _
    ByVal sz As Long _
) As Long



'version/build info
'---------------------------------------------------------------------
'[entry(0x60000023), helpstring("Get libcurl version info")]
'long _stdcall vbcurl_version_info([in] CURLversion age);
Public Declare Function vbcurl_version_info Lib "vblibcurl.dll" (age As CurlVersion) As Long

'[entry(0x6000001d), helpstring("Get the underlying libcurl version string")]
'BSTR _stdcall vbcurl_string_version();
Private Declare Function vbcurl_string_version Lib "vblibcurl.dll" () As Long

'[entry(0x60000027), helpstring("Get supported protocols")]
'void _stdcall vbcurl_version_protocols(
'                [in] long ver,
'                [out] SAFEARRAY(BSTR)* ppsa);
Private Declare Sub vbcurl_version_protocols Lib "vblibcurl.dll" (ByVal hVerInfo As Long, ByRef ary() As Long)


'[entry(0x6000001e), helpstring("Age of libcurl version")]
'long _stdcall vbcurl_version_age([in] long ver);

'[entry(0x6000001f), helpstring("ARES version string")]
'BSTR _stdcall vbcurl_version_ares([in] long ver);

'[entry(0x60000020), helpstring("ARES version number")]
'long _stdcall vbcurl_version_ares_num([in] long ver);

'[entry(0x60000021), helpstring("Bitmask of supported features")]
'long _stdcall vbcurl_version_features([in] long ver);

'[entry(0x60000022), helpstring("Info of host on which libcurl was built")]
'BSTR _stdcall vbcurl_version_host([in] long ver);

'[entry(0x60000024), helpstring("Get libidn version, if present")]
'BSTR _stdcall vbcurl_version_libidn([in] long ver);

'[entry(0x60000025), helpstring("Get libz version, if present")]
'BSTR _stdcall vbcurl_version_libz([in] long ver);
Private Declare Function vbcurl_version_libz Lib "vblibcurl.dll" (ByVal hVer As Long) As Long

'[entry(0x60000026), helpstring("Get numeric version number")]
'long _stdcall vbcurl_version_num([in] long ver);

'[entry(0x60000028), helpstring("Get SSL version string")]
'BSTR _stdcall vbcurl_version_ssl([in] long ver);
Private Declare Function vbcurl_version_ssl Lib "vblibcurl.dll" (ByVal hVer As Long) As Long

'[entry(0x60000029), helpstring("Get SSL version number")]
'long _stdcall vbcurl_version_ssl_num([in] long ver);

'[entry(0x6000002a), helpstring("Get version string")]
'BSTR _stdcall vbcurl_version_string([in] long ver);


Function libCurlVersion() As String
    Dim tmp  As Long, s As String
    tmp = vbcurl_string_version()
    CopyMemory ByVal VarPtr(s), tmp, 4 'steal a ref to an existing BSTR so we now own it
    libCurlVersion = s
End Function

Function libZVersion() As String
    Dim tmp  As Long, s As String, hVer As Long
    hVer = vbcurl_version_info(CURLVERSION_NOW)
    tmp = vbcurl_version_libz(hVer)
    CopyMemory ByVal VarPtr(s), tmp, 4 'steal a ref to an existing BSTR so we now own it
    libZVersion = s
End Function

Function sslVersion() As String
    Dim tmp  As Long, s As String, hVer As Long
    hVer = vbcurl_version_info(CURLVERSION_NOW)
    tmp = vbcurl_version_ssl(hVer)
    CopyMemory ByVal VarPtr(s), tmp, 4 'steal a ref to an existing BSTR so we now own it
    sslVersion = s
End Function

Function libcurlProtocols() As String()
    Dim ptr() As Long, vd As Long, i As Long, s() As String, tmp As String
    
    vd = vbcurl_version_info(CURLVERSION_NOW)
    vbcurl_version_protocols vd, ptr
    
    If AryIsEmpty(ptr) Then Exit Function
    
    ReDim s(UBound(ptr))
    For i = 0 To UBound(ptr)
        CopyMemory ByVal VarPtr(tmp), ptr(i), 4 'steal a ref to an existing BSTR so we now own it
        s(i) = tmp
        'Debug.Print ptr(i) & " " & s(i)
    Next
    
    libcurlProtocols = s()
    
End Function


