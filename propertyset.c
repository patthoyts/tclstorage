/* propertyset.c - Copyright (C) 2005 Pat Thoyts <patthoyts@users.sf.net>
 *
 * Subcommands for manipulating and inspecting property sets within
 * structured storages. 
 *
 * LIMITATIONS
 *   * At this time we only support the standard property sets pre-defined
 *     for COM and Microsoft Office documents.
 *   * The conversion for FILETIME is crap.
 *   * We can currently only set LPSTR values.
 *
 * ----------------------------------------------------------------------
 *
 * See the file "license.terms" for information on usage and redistribution
 * of this file, and for a DISCLAIMER OF ALL WARRANTIES.
 *
 * ----------------------------------------------------------------------
 *
 * @(#) $Id$
 */

#include "tclstorage.h"

typedef struct _PropertySet {
    IPropertyStorage *propPtr;
    FMTID             fmtid;
    DWORD             mode;
} PropertySet;

static long PROPSETID = 0;

       Tcl_ObjCmdProc PropertySetOpenCmd;
       Tcl_ObjCmdProc PropertySetDeleteCmd;
       Tcl_ObjCmdProc PropertySetNamesCmd;

       Tcl_ObjCmdProc PropertyNamesCmd;
static Tcl_ObjCmdProc PropertyGetCmd;
static Tcl_ObjCmdProc PropertySetCmd;
static Tcl_ObjCmdProc PropertyDeleteCmd;
static Tcl_ObjCmdProc PropertyCloseCmd;

static Tcl_CmdDeleteProc PropertyCmdDeleteProc;
static void ConvertValueToString( const PROPVARIANT *propvar, 
                                  WCHAR *pwszValue, ULONG cchValue );

Ensemble PropertyEnsemble[] = {
    { "names",    PropertyNamesCmd,   0 },
    { "get",      PropertyGetCmd,     0 },
    { "set",      PropertySetCmd,     0 },
    { "unset",    PropertyDeleteCmd,   0 },
    { "close",    PropertyCloseCmd,   0 },
    { NULL,       0,                  0 }
};

/*
 * ----------------------------------------------------------------------
 *
 * GetFMTIDFromObj --
 *
 *	Convert a string name into a format identifier that may be used
 *	to access a property set. At the moment we only support the 
 *	standard identifiers - but custom identifiers are possible.
 *
 * Results:
 *	A property set identifier or a Tcl error.
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

static int
GetFMTIDFromObj(Tcl_Interp *interp, Tcl_Obj *objPtr, FMTID *fmtidPtr)
{
    const char *name = Tcl_GetString(objPtr);
    if (strcmp("\005SummaryInformation", name) == 0) {
        memcpy(fmtidPtr, &FMTID_SummaryInformation, sizeof(FMTID));
    } else if (strcmp("\005DocumentSummaryInformation", name) == 0) {
        memcpy(fmtidPtr, &FMTID_DocSummaryInformation, sizeof(FMTID));
    } else if (strcmp("\005UserDefined", name) == 0) {
        memcpy(fmtidPtr, &FMTID_UserDefinedProperties, sizeof(FMTID));
    } else {
        Tcl_Obj *errObj = Tcl_NewStringObj("", -1);
        Tcl_AppendStringsToObj(errObj, "invalid identifier \"", name,
            "\": only \\005SummaryInformation, \\005DocumentSummaryInformation "
            "and \\005UserDefined are supported", (char *)NULL);
        Tcl_SetObjResult(interp, errObj);
        return TCL_ERROR;
    }
    return TCL_OK;
    /*
    typedef struct fmtid_map_type { const char *name; const FMTID fmtid; } fmtid_map_t;
    static const fmtid_map_t map[] = {
        { "\005SummaryInformation", FMTID_SummaryInformation, },
        { "\005DocumentSummaryInformation", FMTID_DocSummaryInformation , },
        { "\005UserDefined", FMTID_UserDefinedProperties , },
        { NULL, 0 }
    };
    int index;
    int r = Tcl_GetIndexFromObjStruct(interp, objPtr, map, sizeof(map[0]), "", 0, &index);
    if (r == TCL_OK) {
        memcpy(fmtidPtr, &map[index].fmtid, sizeof(FMTID));
    }
    return r;
    */
}

/*
 * ----------------------------------------------------------------------
 *
 * GetNameFromPROPID --
 *
 *	Property set items are given integer ids. For the standard
 *	sets these are predefined. This functions returns the name
 *	for a given id in a given property set.
 *
 * Results:
 *	The name for the property id for this property set. NULL if no
 *	name can be provided.
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

const char *summary_names[] = {
    NULL, NULL, "title", "subject", "author", "keywords", "comments",
    "template", "last saved by", "revision number", "total editing time",
    "last printed", "create time", "last saved time", "pages", "words",
     "chars", "thumbnail", "appname", "security"
};

const char *document_names[] = {
    NULL, NULL, "category", "presentation target", "bytes", "lines",
    "paragraphs", "slides", "notes", "hidden slides", "mmclips", "scalecrop",
    "heading pairs", "titles of parts", "manager", "company", "linksuptodate"
};

const char *
GetNameFromPROPID(REFFMTID fmtid, PROPID propid)
{
    const char *name = NULL;
    if (memcmp(&FMTID_SummaryInformation, fmtid, sizeof(FMTID)) == 0) {
        if (propid > 0x01 && propid < 0x14) {
            name = summary_names[(int)propid];
        }
    } else if (memcmp(&FMTID_DocSummaryInformation, fmtid, sizeof(FMTID)) == 0) {
        if (propid > 0x01 && propid < 0x11) {
            name = document_names[(int)propid];
        }
    }
    return name;
}

/*
 * ----------------------------------------------------------------------
 *
 * GetPROPIDFromName --
 *
 *	Convert a string name into a property id. This reverses the 
 *	GetNameFromPROPID function.
 *
 * Results:
 *	A property identifier or 0 if no identifier can be found..
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

static PROPID
GetPROPIDFromName(REFFMTID fmtid, const char *name)
{
    int n;
    if (memcmp(&FMTID_SummaryInformation, fmtid, sizeof(FMTID)) == 0) {
        for (n = 0; n < sizeof(summary_names)/sizeof(summary_names[0]); n++) {
            if (summary_names[n] && stricmp(summary_names[n], name) == 0) {
                return (PROPID)n;
            }
        }
    } else if (memcmp(&FMTID_DocSummaryInformation, fmtid, sizeof(FMTID)) == 0) {
        for (n = 0; n < sizeof(document_names)/sizeof(document_names[0]); n++) {
            if (document_names[n] && stricmp(document_names[n], name) == 0) {
                return (PROPID)n;
            }
        }
    }
    return 0;
}

/*
 * ----------------------------------------------------------------------
 *
 * GetNameFromVarType --
 *
 *	Convert a PROPVARIANT type into a string representation.
 *
 * Results:
 *	Stringify a PROPVARIANT vt type or returns a Tcl error.
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

typedef struct {
    const char *desc;
    VARTYPE vt;
} vt_map_t;
const vt_map_t vt_map[] = {
    { "VT_EMPTY", VT_EMPTY }, {"VT_NULL", VT_NULL}, { "VT_BOOL", VT_BOOL }, { "VT_INT", VT_INT}, 
    { "VT_LPSTR", VT_LPSTR }, { "VT_LPWSTR", VT_LPWSTR }, { "VT_CLSID", VT_CLSID },
    { "VT_FILETIME", VT_FILETIME }, { "VT_DATE", VT_DATE }, {"VT_BSTR", VT_BSTR },
    { NULL, 0 }
};

static int
GetNameFromVarType(Tcl_Interp *interp, VARTYPE vt, Tcl_Obj **namePtrPtr)
{
    const vt_map_t *p;
    for (p = vt_map; p->desc != NULL; p++) {
        if (p->vt == vt) {
            *namePtrPtr = Tcl_NewStringObj(p->desc, -1);
            return TCL_OK;
        }
    }
    Tcl_SetResult(interp, "vt not found", TCL_STATIC);
    return TCL_ERROR;
}

/*
 * ----------------------------------------------------------------------
 *
 * CreatePropertySetCmd --
 *
 *	Creates a Tcl command to use in inspecting and manipulating the
 *	specified property set.
 *
 * Results:
 *	Returns the name of the newly created command.
 *
 * Side effects:
 *	A new Tcl command is registered in the current interpreter.
 */

static int
CreatePropertySetCmd(Tcl_Interp *interp, FMTID fmtid, IPropertyStorage *propPtr, DWORD mode)
{
    EnsembleCmdData *dataPtr;
    PropertySet *propsetPtr;
    Tcl_Obj *nameObj = NULL;
    char name[7 + TCL_INTEGER_SPACE];
    long id = InterlockedIncrement(&PROPSETID);
		
    _snprintf(name, 7 + TCL_INTEGER_SPACE, "propset%lu", id);
    nameObj = Tcl_NewStringObj(name, -1);
    
    propsetPtr = (PropertySet *)ckalloc(sizeof(PropertySet));
    propsetPtr->mode = mode;
    memcpy(&propsetPtr->fmtid, &fmtid, sizeof(FMTID));
    propsetPtr->propPtr = propPtr;
    propsetPtr->propPtr->lpVtbl->AddRef(propsetPtr->propPtr);
		
    dataPtr = (EnsembleCmdData *)ckalloc(sizeof(EnsembleCmdData));
    dataPtr->ensemble = PropertyEnsemble;
    dataPtr->clientData = propsetPtr;
    Tcl_CreateObjCommand(interp, name, TclEnsembleCmd,
        (ClientData)dataPtr,
        (Tcl_CmdDeleteProc *)PropertyCmdDeleteProc);
		
    Tcl_SetObjResult(interp, nameObj);
    return TCL_OK;
}

/*
 * ----------------------------------------------------------------------
 *
 * PropertyCmdDeleteProc -
 *
 *	Clean up the allocated memory associated with the property set
 *	command.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	Memory free'd, COM objects released.
 *
 * ----------------------------------------------------------------------
 */

static void
PropertyCmdDeleteProc(ClientData clientData)
{
    EnsembleCmdData *dataPtr = (EnsembleCmdData *)clientData;
    PropertySet *propsetPtr = (PropertySet *)dataPtr->clientData;
    propsetPtr->propPtr->lpVtbl->Commit(propsetPtr->propPtr, STGC_DEFAULT);
    propsetPtr->propPtr->lpVtbl->Release(propsetPtr->propPtr);
    ckfree((char *)propsetPtr);
    ckfree((char *)dataPtr);
}

/*
 * ----------------------------------------------------------------------
 *
 * PropertySetOpenCmd --
 *
 *	Opens a property set and creates a Tcl command to use in accessing
 *	the properties in the set. Access permissions depend upon the 
 *	access permissions for the open storage as well as those specified
 *	here.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	A new Tcl command is created in the current interpreter.
 *
 * ----------------------------------------------------------------------
 */

int
PropertySetOpenCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    int r = TCL_OK;
    
    if (objc < 4 || objc > 5) {

        Tcl_WrongNumArgs(interp, 3, objv, "id ?mode?");
        r = TCL_ERROR;

    } else {
        
        IStorage *stgPtr = storagePtr->pstg;
        IPropertySetStorage *setPtr;
        HRESULT hr = S_OK;
        DWORD grfMode =  STGM_DIRECT | STGM_SHARE_EXCLUSIVE;
        FMTID fmtid;
        
        if (GetFMTIDFromObj(interp, objv[3], &fmtid) != TCL_OK)
            return TCL_ERROR;

        if (objc > 4) {
            r = GetStorageFlagsFromObj(interp, objv[4], &grfMode);
        } else {
            r |= STGM_READ;
        }

        hr = stgPtr->lpVtbl->QueryInterface(stgPtr, &IID_IPropertySetStorage, (void**)&setPtr);
        if (SUCCEEDED(hr)) {
	        IPropertyStorage *propPtr;
            if (grfMode & STGM_CREATE)
                hr = setPtr->lpVtbl->Create(setPtr, &fmtid, &fmtid, PROPSETFLAG_DEFAULT, grfMode & STGM_WIN32MASK, &propPtr);
            else
                hr = setPtr->lpVtbl->Open(setPtr, &fmtid, grfMode & STGM_WIN32MASK, &propPtr);
            if (SUCCEEDED(hr)) {
                r = CreatePropertySetCmd(interp, fmtid, propPtr, grfMode);
                propPtr->lpVtbl->Release(propPtr);
            }
            setPtr->lpVtbl->Release(setPtr);
        }
        if (FAILED(hr)) {
            Tcl_Obj *errObj = Tcl_NewStringObj("", 0);
            Tcl_AppendStringsToObj(errObj, "error opening property set \"",
                Tcl_GetString(objv[3]), "\":", (char *)NULL);
            Tcl_AppendObjToObj(errObj, Win32Error("", hr));
            Tcl_SetObjResult(interp, errObj);
            r = TCL_ERROR;
        }
    }
        
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * PropertySetDeleteCmd --
 *
 *	Delete the specified property set.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	The property set is removed from the storage.
 *
 * ----------------------------------------------------------------------
 */

int
PropertySetDeleteCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    int r = TCL_OK;
    
    if (objc != 4) {

        Tcl_WrongNumArgs(interp, 3, objv, "id");
        r = TCL_ERROR;

    } else {
        
        IStorage *stgPtr = storagePtr->pstg;
        IPropertySetStorage *setPtr;
        
        HRESULT hr = stgPtr->lpVtbl->QueryInterface(stgPtr, &IID_IPropertySetStorage, (void**)&setPtr);
        if (SUCCEEDED(hr)) {

            Tcl_SetResult(interp, "error: command not implemented", TCL_STATIC);
            hr = E_FAIL;
            setPtr->lpVtbl->Release(setPtr);
        }

        if (FAILED(hr))
            Tcl_SetObjResult(interp, Win32Error("error", hr));
        r = SUCCEEDED(hr) ? TCL_OK : TCL_ERROR;
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * PropertySetNamesCmd --
 *
 *	Enumerate all available property sets in the current storage.
 *
 * Results:
 *	A standard Tcl result. The interpreter result is set to a Tcl
 *	list containing the property set names.
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

int
PropertySetNamesCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    int r = TCL_OK;
    
    if (objc != 3) {

        Tcl_WrongNumArgs(interp, 3, objv, "");
        r = TCL_ERROR;

    } else {
        
        IStorage *stgPtr = storagePtr->pstg;
        IPropertySetStorage *setPtr;
        
        HRESULT hr = stgPtr->lpVtbl->QueryInterface(stgPtr, &IID_IPropertySetStorage, (void**)&setPtr);
        if (SUCCEEDED(hr)) {
            IEnumSTATPROPSETSTG *enumPtr = NULL;
            hr = setPtr->lpVtbl->Enum(setPtr, &enumPtr);
            if (SUCCEEDED(hr)) {
                int n, nret = 0;
                STATPROPSETSTG astat[12];
                Tcl_Obj *retObj = Tcl_NewListObj(0, NULL);
                do {
                    hr = enumPtr->lpVtbl->Next(enumPtr, 12, astat, &nret);
                    for (n = 0; n < nret; n++) {
                        WCHAR wsz[64];
                        StringFromGUID2(&astat[n].fmtid, wsz, sizeof(wsz)/sizeof(wsz[0]));
                        Tcl_ListObjAppendElement(interp, retObj, Tcl_NewUnicodeObj(wsz, -1));
                    }
                } while (hr == S_OK);
                if (SUCCEEDED(hr)) {
                    Tcl_SetObjResult(interp, retObj);
                }
                enumPtr->lpVtbl->Release(enumPtr);
            }
            setPtr->lpVtbl->Release(setPtr);
        }
        if (FAILED(hr))
            Tcl_SetObjResult(interp, Win32Error("error", hr));
        r = SUCCEEDED(hr) ? TCL_OK : TCL_ERROR;
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * PropertyNamesCmd --
 *
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

int
PropertyNamesCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    PropertySet *setPtr = (PropertySet *)clientData;
    IEnumSTATPROPSTG *enumPtr;

    HRESULT hr = setPtr->propPtr->lpVtbl->Enum(setPtr->propPtr, &enumPtr);
    if (SUCCEEDED(hr)) {
	STATPROPSTG astat[12];
	ULONG nret, n;
    const char *propname = NULL;
	Tcl_Obj *resObj = Tcl_NewListObj(0, NULL);
	do {
	    hr = enumPtr->lpVtbl->Next(enumPtr, 12, astat, &nret);
	    for (n = 0; n < nret; n++) {
		Tcl_Obj *vtObj;
        
        if (astat[n].lpwstrName != NULL) {
		    Tcl_ListObjAppendElement(interp, resObj, Tcl_NewUnicodeObj(astat[n].lpwstrName, -1));
        } else {
            if ((propname = GetNameFromPROPID(&setPtr->fmtid, astat[n].propid)) == NULL) {
		        Tcl_ListObjAppendElement(interp, resObj, Tcl_NewLongObj(astat[n].propid));
            } else {
                Tcl_ListObjAppendElement(interp, resObj, Tcl_NewStringObj(propname, -1));
            }
        }

		if (GetNameFromVarType(interp, astat[n].vt, &vtObj) == TCL_OK) {
		    Tcl_ListObjAppendElement(interp, resObj, vtObj);
		} else {
		    Tcl_ListObjAppendElement(interp, resObj, Tcl_NewLongObj(astat[n].vt));
		}

		CoTaskMemFree(astat[n].lpwstrName);
	    }
	} while (hr == S_OK);
	
	enumPtr->lpVtbl->Release(enumPtr);
	Tcl_SetObjResult(interp, resObj);
    }
    if (FAILED(hr)) {
	Tcl_SetObjResult(interp, Win32Error("error", hr));
    }
    return SUCCEEDED(hr) ? TCL_OK : TCL_ERROR;
}

/*
 * ----------------------------------------------------------------------
 *
 * PropertyGetCmd --
 *
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

int
PropertyGetCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    PropertySet *setPtr = (PropertySet *)clientData;
    PROPSPEC spec;
    PROPVARIANT v;
    HRESULT hr = S_OK;

    if (objc != 3) {
        Tcl_WrongNumArgs(interp, 2, objv, "name");
        return TCL_ERROR;
    }

    PropVariantInit(&v);

    spec.ulKind = PRSPEC_PROPID;
    spec.propid = GetPROPIDFromName(&setPtr->fmtid, Tcl_GetString(objv[2]));
    if (spec.propid == 0) {
        spec.ulKind = PRSPEC_LPWSTR;
        spec.lpwstr = Tcl_GetUnicode(objv[2]);
    }
    hr = setPtr->propPtr->lpVtbl->ReadMultiple(setPtr->propPtr, 1, &spec, &v);
    if (SUCCEEDED(hr)) {
        WCHAR wsz[1024];
        ConvertValueToString(&v, wsz, 1024);
        PropVariantClear(&v);
        Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(wsz, -1));
    }
    if (FAILED(hr))
        Tcl_SetObjResult(interp, Win32Error("error", hr));
    return SUCCEEDED(hr) ? TCL_OK : TCL_ERROR;
}

/*
 * ----------------------------------------------------------------------
 *
 * PropertySetCmd --
 *
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

int
PropertySetCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    PropertySet *setPtr = (PropertySet *)clientData;
    PROPSPEC spec;
    PROPVARIANT v;
    HRESULT hr = S_OK;

    if (objc < 4 || objc > 5) {
        Tcl_WrongNumArgs(interp, 2, objv, "name value ?type?");
        return TCL_ERROR;
    }

    PropVariantInit(&v);

    spec.ulKind = PRSPEC_PROPID;
    spec.propid = GetPROPIDFromName(&setPtr->fmtid, Tcl_GetString(objv[2]));
    if (spec.propid == 0) {
        spec.ulKind = PRSPEC_LPWSTR;
        spec.lpwstr = Tcl_GetUnicode(objv[2]);
    }

    v.vt = VT_LPSTR;
    v.pszVal = Tcl_GetString(objv[3]);

    hr = setPtr->propPtr->lpVtbl->WriteMultiple(setPtr->propPtr, 1, &spec, &v, 2);
    /* PropVariantClear(&v); */
    if (FAILED(hr))
        Tcl_SetObjResult(interp, Win32Error("error", hr));
    return SUCCEEDED(hr) ? TCL_OK : TCL_ERROR;
}

/*
 * ----------------------------------------------------------------------
 *
 * PropertyDeleteCmd --
 *
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

int
PropertyDeleteCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    PropertySet *setPtr = (PropertySet *)clientData;
    PROPSPEC spec;
    HRESULT hr = S_OK;

    if (objc != 3) {
        Tcl_WrongNumArgs(interp, 2, objv, "name");
        return TCL_ERROR;
    }

    spec.ulKind = PRSPEC_PROPID;
    spec.propid = GetPROPIDFromName(&setPtr->fmtid, Tcl_GetString(objv[2]));
    if (spec.propid == 0) {
        spec.ulKind = PRSPEC_LPWSTR;
        spec.lpwstr = Tcl_GetUnicode(objv[2]);
    }

    hr = setPtr->propPtr->lpVtbl->DeleteMultiple(setPtr->propPtr, 1, &spec);
    if (FAILED(hr))
        Tcl_SetObjResult(interp, Win32Error("error", hr));
    return SUCCEEDED(hr) ? TCL_OK : TCL_ERROR;
}

/*
 * ----------------------------------------------------------------------
 *
 * PropertyCloseCmd --
 *
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

int
PropertyCloseCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    PropertySet *setPtr = (PropertySet *)clientData;
    if (objc != 2) {
        Tcl_WrongNumArgs(interp, 2, objv, "");
        return TCL_ERROR;
    }
    Tcl_DeleteCommand(interp, Tcl_GetString(objv[0]));
    return TCL_OK;
}

/*
 * ----------------------------------------------------------------------
 *
 * ConvertValueToString --
 *
 *	Convert a PROPVARIANT into a string value. This is modified from
 *	a Microsoft sample.
 *
 * Results:
 *	The pwszValue value is set to the string representation of the 
 *	variant data.
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

void
ConvertValueToString( const PROPVARIANT *propvar,
                      WCHAR *pwszValue,
                      ULONG cchValue )
{
    pwszValue[ cchValue - 1 ] = L'\0';
    --cchValue;

    switch( propvar->vt )
    {
    case VT_EMPTY:
        wcsncpy( pwszValue, L"", cchValue );
        break;
    case VT_NULL:
        wcsncpy( pwszValue, L"", cchValue );
        break;
    case VT_I2:
        _snwprintf( pwszValue, cchValue, L"%i", propvar->iVal );
        break;
    case VT_I4:
    case VT_INT:
        _snwprintf( pwszValue, cchValue, L"%li", propvar->lVal );
        break;
    case VT_I8:
        _snwprintf( pwszValue, cchValue, L"%I64i", propvar->hVal );
        break;
    case VT_UI2:
        _snwprintf ( pwszValue, cchValue, L"%u", propvar->uiVal );
        break;
    case VT_UI4:
    case VT_UINT:
        _snwprintf ( pwszValue, cchValue, L"%lu", propvar->ulVal );
        break;
    case VT_UI8:
        _snwprintf ( pwszValue, cchValue, L"%I64u", propvar->uhVal );
        break;
    case VT_R4:
        _snwprintf ( pwszValue, cchValue, L"%f", propvar->fltVal );
        break;
    case VT_R8:
        _snwprintf ( pwszValue, cchValue, L"%lf", propvar->dblVal );
        break;
    case VT_BSTR:
        _snwprintf ( pwszValue, cchValue, L"%s", propvar->bstrVal );
        break;
    case VT_ERROR:
        _snwprintf ( pwszValue, cchValue, L"0x%08X", propvar->scode );
        break;
    case VT_BOOL:
        _snwprintf ( pwszValue, cchValue, L"%s",
                     VARIANT_TRUE == propvar->boolVal ? L"true" : L"false" );
        break;
    case VT_I1:
        _snwprintf ( pwszValue, cchValue, L"%i", propvar->cVal );
        break;
    case VT_UI1:
        _snwprintf ( pwszValue, cchValue, L"%u", propvar->bVal );
        break;
    case VT_VOID:
        wcsncpy( pwszValue, L"", cchValue );
        break;
    case VT_LPSTR:
        if( 0 >_snwprintf ( pwszValue, cchValue, L"%hs", propvar->pszVal ))
            wcsncpy( pwszValue, L"...", cchValue );
        break;
    case VT_LPWSTR:
        if( 0 > _snwprintf ( pwszValue, cchValue, L"%s", propvar->pwszVal ))
            wcsncpy( pwszValue, L"...", cchValue );
        break;
    case VT_FILETIME:
        _snwprintf ( pwszValue, cchValue, L"%08x:%08x",
                     propvar->filetime.dwHighDateTime,
                     propvar->filetime.dwLowDateTime );
        break;
    case VT_CLSID:
        pwszValue[0] = L'\0';
        StringFromGUID2( propvar->puuid, pwszValue, cchValue );
        break;
    default:
        wcsncpy( pwszValue, L"...", cchValue );
        break;
    }
}

/* ----------------------------------------------------------------------
 *
 * Local variables:
 * mode: c
 * indent-tabs-mode: nil
 * End:
 */
