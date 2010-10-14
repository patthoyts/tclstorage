/* tclstorage.c - Copyright (C) 2004 Pat Thoyts <patthoyts@users.sf.net>
 *
 * This file is Tcl extension that adds a 'storage' command to Tcl
 * and provides access to Microsoft's "Structured Storage" file format.
 * Structured storages are used extensively to provide persistence for
 * OLE or COM components. The format presents a filesystem-like
 * hierarchy of storages and streams that maps well into Tcl's 
 * virtual filesystem model.
 * 
 * Notable users of structured storages are Microsoft Word and Excel.
 *
 * Usage:
 *   storage open filename mode
 *      mode is as per the Tcl open command "[raw]+?"
 *      returns a storage command. The storage will remain open
 *      as long as the command exists. You can close the storage file
 *      using either the close subcommand or renaming the command.
 *   eg: % storage open document.doc r+
 *       stg1
 *
 *  object commands:
 *   opendir name ?mode?     open or create a sub-storage
 *   open name ?mode?        open or create a stream as a Tcl channel
 *   close                   close the storage or sub-storage
 *   stat name varname       get information about the named item
 *   commit                  not used
 *   rename oldname newname  rename a stream or sub-storage
 *   remove name             deletes a stream or sub-storage + contents
 *   names                   list all items in the current storage
 *   propertyset             subcommands to handle property sets
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


#ifndef PACKAGE_NAME
#define PACKAGE_NAME       "Storage"
#endif
#ifndef PACKAGE_VERSION
#define PACKAGE_VERSION    "1.3.0"
#endif

#include "tclstorage.h"

#if _MSC_VER >= 1000
#pragma comment(lib, "ole32")
#pragma comment(lib, "advapi32")
#endif

static Tcl_ObjCmdProc StorageCmd;
static Tcl_CmdDeleteProc StorageCmdDeleteProc;
static Tcl_ObjCmdProc StorageObjCmd;
static Tcl_CmdDeleteProc StorageObjDeleteProc;
static Tcl_ObjCmdProc StorageOpendirCmd;
static Tcl_ObjCmdProc StorageOpenCmd;
static Tcl_ObjCmdProc StorageStatCmd;
static Tcl_ObjCmdProc StorageRenameCmd;
static Tcl_ObjCmdProc StorageRemoveCmd;
static Tcl_ObjCmdProc StorageCloseCmd;
static Tcl_ObjCmdProc StorageCommitCmd;
static Tcl_ObjCmdProc StorageNamesCmd;

extern Tcl_ObjCmdProc PropertySetOpenCmd;
extern Tcl_ObjCmdProc PropertySetDeleteCmd;
extern Tcl_ObjCmdProc PropertySetNamesCmd;

static long UNIQUEID = 0;

static Tcl_InterpDeleteProc PackageDeleteProc;
static int GetItemInfo(Tcl_Interp *interp, IStorage *pstg, 
    Tcl_Obj *pathObj, STATSTG *pstatstg);
static void TimeToFileTime(time_t t, LPFILETIME pft);
static time_t TimeFromFileTime(const FILETIME *pft);


static Tcl_DriverCloseProc     StorageChannelClose;
static Tcl_DriverInputProc     StorageChannelInput;
static Tcl_DriverOutputProc    StorageChannelOutput;
static Tcl_DriverSeekProc      StorageChannelSeek;
static Tcl_DriverWatchProc     StorageChannelWatch;
static Tcl_DriverGetHandleProc StorageChannelGetHandle;
static Tcl_DriverWideSeekProc  StorageChannelWideSeek;

static int EventProc(Tcl_Event *evPtr, int flags);
static void SetupProc(ClientData clientData, int flags);
static void CheckProc(ClientData clientData, int flags);

#define STORAGE_PACKAGE_KEY  "StoragePackageKey"
#define STORAGE_FLAG_ASYNC   (1<<1)
#define STORAGE_FLAG_PENDING (1<<2)

struct Package;

typedef struct StorageChannel {
    Tcl_Channel chan;
    struct Package *pkgPtr;
    struct StorageChannel *nextPtr;
    Tcl_Interp *interp;
    DWORD grfMode;
    int watchmask;
    int validmask;
    int flags;
    IStream *pstm;
} StorageChannel;

typedef struct Package {
    struct StorageChannel *headPtr;
    unsigned long count;
    unsigned long uid;
} Package;

typedef struct ChannelEvent {
    Tcl_Event header;
    StorageChannel *instPtr;
    int flags;
} ChannelEvent;

static Tcl_ChannelType StorageChannelType = {
    "storage",
    (Tcl_ChannelTypeVersion)TCL_CHANNEL_VERSION_2,
    StorageChannelClose,
    StorageChannelInput,
    StorageChannelOutput,
    StorageChannelSeek,
    /* StorageChannelSetOptions */ NULL,
    /* StorageChannelGetOptions */ NULL,
    StorageChannelWatch,
    StorageChannelGetHandle,
    /* StorageChannelClose2 */     NULL,
    /* StorageChannelBlockMode */  NULL,
    /* StorageChannelFlush */      NULL,
    /* StorageChannelHandler */    NULL,
    StorageChannelWideSeek
};


static Ensemble StorageEnsemble[] = {
    { "open",   Storage_OpenStorage,   0 },
    { NULL,     0,                     0 }
};

static Ensemble PropertySetEnsemble[] = {
    { "open",     PropertySetOpenCmd,    0 },
    { "delete",   PropertySetDeleteCmd,  0 },
    { "names",    PropertySetNamesCmd,   0 },
    { NULL,       0,                     0 },
};

static Ensemble StorageObjEnsemble[] = {
    { "opendir",     StorageOpendirCmd,     0 },
    { "open",        StorageOpenCmd,        0 },
    { "close",       StorageCloseCmd,       0 },
    { "stat",        StorageStatCmd,        0 },
    { "commit",      StorageCommitCmd,      0 },
    { "rename",      StorageRenameCmd,      0 },
    { "remove",      StorageRemoveCmd,      0 },
    { "names",       StorageNamesCmd,       0 },
    { "propertyset", NULL, PropertySetEnsemble},
    { NULL,          0,                     0 }
};

/* ---------------------------------------------------------------------- */

typedef struct {
    const char *s;
    const int   posixmode;
    const DWORD f;
} stgm_map_t;
const stgm_map_t stgm_map[] = {
    { "r",  0x01, STGM_READ },
    { "r+", 0x05, STGM_READWRITE },
    { "w",  0x12, STGM_WRITE|STGM_CREATE },
    { "w+", 0x16, STGM_READWRITE|STGM_CREATE },
    { "a",  0x02, STGM_WRITE|STGM_APPEND },
    { "a+", 0x06, STGM_READWRITE|STGM_APPEND },
    { NULL, 0}
};

/*
 * ----------------------------------------------------------------------
 *
 * GetStorageFlagsFromObj --
 *
 *	Converts a mode string as documented for the 'open' command
 *	into a set of STGM enumeration flags for use with the 
 *	storage implementaiton.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	The location contained in the flags pointer will have bits set.
 *
 * ----------------------------------------------------------------------
 */

int 
GetStorageFlagsFromObj(Tcl_Interp *interp, 
    Tcl_Obj *objPtr, int *flagsPtr)
{
    int index = 0, objc, n, r = TCL_OK;
    Tcl_Obj **objv;
    
    r = Tcl_ListObjGetElements(interp, objPtr, &objc, &objv);
    if (r == TCL_OK) {
        for (n = 0; n < objc; n++) {
            r = Tcl_GetIndexFromObjStruct(interp, objv[n],
		stgm_map, sizeof(stgm_map[0]), "storage flag", 0, &index);
            if (r == TCL_OK)
                *flagsPtr |= stgm_map[index].f;
        }
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * Storage_Init --
 *
 *	Initialize the Storage package.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	The Storage package is provided.
 *	One new command 'storage' is added to the current interpreter.
 *
 * ----------------------------------------------------------------------
 */

int
Storage_Init(Tcl_Interp *interp)
{
    EnsembleCmdData *dataPtr;
    Package *pkgPtr;

    if (Tcl_InitStubs(interp, "8.2", 0) == NULL) {
        return TCL_ERROR;
    }
    
    pkgPtr = (Package *)ckalloc(sizeof(Package));
    pkgPtr->headPtr = NULL;
    pkgPtr->count = 0;
    pkgPtr->uid = 0;
    Tcl_CreateEventSource(SetupProc, CheckProc, pkgPtr);
    Tcl_SetAssocData(interp, STORAGE_PACKAGE_KEY, PackageDeleteProc, pkgPtr);

    dataPtr = (EnsembleCmdData *)ckalloc(sizeof(EnsembleCmdData));
    dataPtr->ensemble = StorageEnsemble;
    dataPtr->clientData = NULL;
    Tcl_CreateObjCommand(interp, "storage", TclEnsembleCmd, 
	(ClientData)dataPtr,
	(Tcl_CmdDeleteProc *)StorageCmdDeleteProc);
    return Tcl_PkgProvide(interp, PACKAGE_NAME, PACKAGE_VERSION);
}

/*
 * ----------------------------------------------------------------------
 *
 * Storage_SafeInit -
 *
 *	Initialize the package in a safe interpreter.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	See Storage_Init.
 *
 * ----------------------------------------------------------------------
 */

int
Storage_SafeInit(Tcl_Interp *interp)
{
    return Storage_Init(interp);
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageCmdDeleteProc -
 *
 *	Clean up the allocated memory associated with the storage command.
 *
 * Results:
 *	None.
 *
 * Side effects:
 *	Memory free'd
 *
 * ----------------------------------------------------------------------
 */

static void
StorageCmdDeleteProc(ClientData clientData)
{
    EnsembleCmdData *data = (EnsembleCmdData *)clientData;
    ckfree((char *)data);
}

/*
 * ----------------------------------------------------------------------
 *
 * PackageDeleteProc -
 *
 *	Clean up the allocated memory associated with the package.
 *
 * Results:
 *	None.
 *
 * Side effects:
 *	Memory free'd and the event source removed.
 *
 * ----------------------------------------------------------------------
 */

static void
PackageDeleteProc(ClientData clientData, Tcl_Interp *interp)
{
    Package *pkgPtr = clientData;
    Tcl_DeleteEventSource(SetupProc, CheckProc, pkgPtr);
    ckfree((char *)pkgPtr);
}

/*
 * ----------------------------------------------------------------------
 *
 * CreateStorageCommand -
 *
 *	Utility function to create a unique Tcl command to represent
 *	a Structured storage instance.
 *
 * Results:
 *	A standard Tcl result. The name of the new command is returned
 *	as the interp result.
 *
 * Side effects:
 *	A new command is created in the Tcl interpreter.
 *	The command name is added to a list held by the parent storage.
 *
 * ----------------------------------------------------------------------
 */

static int
CreateStorageCommand(Tcl_Interp *interp, Storage *parentPtr, 
    IStorage *pstg, int mode)
{
    EnsembleCmdData *dataPtr = NULL;
    Storage *storagePtr = NULL;
    Tcl_Obj *nameObj = NULL;
    char name[3 + TCL_INTEGER_SPACE];
    long id = InterlockedIncrement(&UNIQUEID);
    
    _snprintf(name, 3 + TCL_INTEGER_SPACE, "stg%lu", id);
    nameObj = Tcl_NewStringObj(name, -1);
    
    dataPtr = (EnsembleCmdData *)ckalloc(sizeof(EnsembleCmdData));
    storagePtr = (Storage *)ckalloc(sizeof(Storage));
    storagePtr->mode = mode;
    storagePtr->pstg = pstg;
    storagePtr->children = Tcl_NewListObj(0, NULL);
    
    Tcl_IncrRefCount(storagePtr->children);
    
    dataPtr->clientData = storagePtr;
    dataPtr->ensemble = StorageObjEnsemble;
    
    Tcl_CreateObjCommand(interp, name, TclEnsembleCmd, 
	(ClientData)dataPtr, (Tcl_CmdDeleteProc *)StorageObjDeleteProc);
    
    if (parentPtr) {
        Tcl_ListObjAppendElement(interp, parentPtr->children, nameObj);
    }
    
    Tcl_SetObjResult(interp, nameObj);
    return TCL_OK;
}

/*
 * ----------------------------------------------------------------------
 *
 * Storage_OpenStorage -
 *
 *	Creates or opens a structured storage file. This will create 
 *	a unique command in the Tcl interpreter that can be used to 
 *	access the contents of the storage. The file will remain
 *	open with exclusive access until this command is destroyed either
 *	by the use of the close sub-command or by renaming the command
 *	to {}.
 *	The mode string is as per the Tcl open command. If w is specified
 *	the file will be created.
 *
 * Results:
 *	A standard Tcl result. The name of the new command is placed in
 *	the interpreters result.
 *
 * Side effects:
 *	The named storage is opened exclusively according to the mode 
 *	 given and a new Tcl command is created.
 *
 * ----------------------------------------------------------------------
 */

int
Storage_OpenStorage(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    HRESULT hr = S_OK;
    int r = TCL_OK;
    int mode = STGM_DIRECT | STGM_SHARE_EXCLUSIVE;
    IStorage *pstg = NULL;
    
    if (objc < 3 || objc > 4) {
        Tcl_WrongNumArgs(interp, 2, objv, "filename ?access?");
        return TCL_ERROR;
    }
    if (objc == 4) {
        r = GetStorageFlagsFromObj(interp, objv[3], &mode);
    } else {
        mode |= STGM_READ;
    }
    
    if (r == TCL_OK) {
        if (mode & STGM_CREATE) {
	    int cchFile = 0;
	    LPCWSTR wszFile = Tcl_GetUnicodeFromObj(objv[2], &cchFile);
	    if (cchFile < 1) {
		ILockBytes *pLockBytes = NULL;
		hr = CreateILockBytesOnHGlobal(NULL, TRUE, &pLockBytes);
		if (SUCCEEDED(hr)) {
		    hr = StgCreateDocfileOnILockBytes(pLockBytes,
			mode & STGM_WIN32MASK, 0, &pstg);
		    pLockBytes->lpVtbl->Release(pLockBytes);
		}
	    } else {
		hr = StgCreateDocfile(Tcl_GetUnicode(objv[2]), 
		    mode & STGM_WIN32MASK, 0, &pstg);
	    }
	} else {
            hr = StgOpenStorage(Tcl_GetUnicode(objv[2]), NULL, 
		mode & STGM_WIN32MASK, NULL, 0, &pstg);
	}
	
        if (SUCCEEDED(hr)) {
            r = CreateStorageCommand(interp, NULL, pstg, mode);
        } else {
            Tcl_Obj *errObj = Win32Error("failed to open storage", hr);
            Tcl_SetObjResult(interp, errObj);
            r = TCL_ERROR;
        }
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageObjDeleteProc -
 *
 *	Callback to handle storage object command deletion.
 *
 * Results:
 *	None.
 *
 * Side effects:
 *	Allocated resources are free'd and the IStorage pointer is
 *	released which frees COM resources. This also unlocks the 
 *	associated file.
 *
 * ----------------------------------------------------------------------
 */

static void
StorageObjDeleteProc(ClientData clientData)
{
    EnsembleCmdData *dataPtr = (EnsembleCmdData *)clientData;
    Storage *storagePtr = (Storage *)dataPtr->clientData;
    
    if (storagePtr->pstg)
        storagePtr->pstg->lpVtbl->Release(storagePtr->pstg);
    Tcl_DecrRefCount(storagePtr->children);
    ckfree((char *)storagePtr);
    ckfree((char *)dataPtr);
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageCloseCmd -
 *
 *	Closes the storage instance and deletes the command.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	See StorageObjDeleteProc
 *
 * ----------------------------------------------------------------------
 */

static int
StorageCloseCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    IStorage *pstg = storagePtr->pstg;
    int r = TCL_OK;
    
    if (objc > 2) {
        Tcl_WrongNumArgs(interp, 2, objv, "");
        r = TCL_ERROR;
    } else {
        /* We may need to delete all child storages too, because they
         * will become unusable anyway. Alternatively we could refuse
         * to close this one because it has children?  At the moment
         * the tcl code in the vfs package does this for us.  
         */
        Tcl_DeleteCommand(interp, Tcl_GetString(objv[0]));
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageCommitCmd -
 *
 *	Flush changes to the underlying file.
 *	At the moment we always use STGM_DIRECT. In the future we may
 *	support transacted mode in which case this would do something.
 *	However, for multimegabyte files there is a _significant_
 *	performance hit when using transacted mode - especially during
 *	the commit.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	None. In the future this may flush unsaved changes to the file.
 *
 * ----------------------------------------------------------------------
 */

static int
StorageCommitCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    IStorage *pstg = storagePtr->pstg;
    HRESULT hr = S_OK;
    int r = TCL_OK;
    
    if (objc > 2) {
        Tcl_WrongNumArgs(interp, 2, objv, "");
        r = TCL_ERROR;
    } else {
        hr = pstg->lpVtbl->Commit(pstg, 0);
        if (FAILED(hr)) {
            Tcl_SetObjResult(interp, Win32Error("commit error", hr));
            r = TCL_ERROR;
        }
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageOpendirCmd -
 *
 *	Opens a sub-storage. A new Tcl command is created to manage the
 *	resource and the mode is as per the Tcl open command. If 'w'
 *	is specified then the sub-storage is created as a child of the
 *	current storage if it is not already present.
 *	Note: Storages may be read-only or write-only or read-write.
 *
 *	The sub-storage is only usable if all it's parents are still
 *	open. This limitation is part of the COM architecture. 
 *	If a parent storage is closed then the only valid command
 *	on its children is a close.
 *
 * Results:
 *	A standard Tcl result. The name of the new command is placed
 *	in the interpreter's result.
 *
 * Side effects:
 *	A new command is created in the Tcl interpreter and associated
 *	with the sub-storage. 
 *
 * ----------------------------------------------------------------------
 */

static int
StorageOpendirCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    IStorage *pstg = storagePtr->pstg;
    IStorage *pstgNew = NULL;
    HRESULT hr = S_OK;
    int mode = storagePtr->mode;
    int r = TCL_OK;
    
    
    if (objc < 3 || objc > 4) {
        Tcl_WrongNumArgs(interp, 2, objv, "dirname mode");
        return TCL_ERROR;
    }
    
    if (objc == 4) {
        mode &= STGM_STREAMMASK;
        r = GetStorageFlagsFromObj(interp, objv[3], &mode);
    } else {
        mode &= ~STGM_CREATE;
    }
    
    hr = pstg->lpVtbl->OpenStorage(pstg, Tcl_GetUnicode(objv[2]), NULL,
        (mode & ~STGM_CREATE) & STGM_WIN32MASK, NULL, 0, &pstgNew);
    if (FAILED(hr)) {
        if (mode & STGM_CREATE) {
            hr = pstg->lpVtbl->CreateStorage(pstg, Tcl_GetUnicode(objv[2]), 
                mode & STGM_WIN32MASK, 0, 0, &pstgNew);
        }
        if (FAILED(hr)) {
            Tcl_Obj *errObj = Tcl_NewStringObj("", 0);
            Tcl_AppendStringsToObj(errObj, "could not ", 
                (mode & STGM_CREATE) ? "create" : "open",
                " \"", Tcl_GetString(objv[2]), "\"", (char *)NULL);
            Tcl_AppendObjToObj(errObj, Win32Error("", hr));
            Tcl_SetObjResult(interp, errObj);
            r = TCL_ERROR;
        }
    }
    if (SUCCEEDED(hr)) {
        r = CreateStorageCommand(interp, storagePtr, pstgNew, mode);
    }
    
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageOpenCmd -
 *
 *	Open a file within the storage. This opens the named item
 *	and creates a Tcl channel to support reading and writing
 *	data. Modes are as per the Tcl 'open' command and may depend upon
 *	the mode settings of the owning storage.
 *
 * Results:
 *	A standard Tcl result. The channel name is returned in the Tcl
 *	interpreter result.
 *
 * Side effects:
 *	A new stream may be created with the given name.
 *	A Tcl channel is created in the Tcl interpreter.
 *
 * ----------------------------------------------------------------------
 */

static int
StorageOpenCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    IStorage *pstg = storagePtr->pstg;
    IStream *pstm = NULL;
    int r = TCL_OK;
    int mode = storagePtr->mode;
    
    if (objc < 3 || objc > 4) {
        Tcl_WrongNumArgs(interp, 2, objv, "filename mode");
        return TCL_ERROR;
    }
    if (objc == 4) {
        mode &= STGM_STREAMMASK; 
        r = GetStorageFlagsFromObj(interp, objv[3], &mode);
    } else {
        mode &= STGM_STREAMMASK;
        mode |= STGM_READ;
    }
    
    if (r == TCL_OK) {
	
        HRESULT hr = S_OK;
        if (mode & STGM_CREATE) {
            hr = pstg->lpVtbl->CreateStream(pstg, Tcl_GetUnicode(objv[2]),
		mode & STGM_WIN32MASK,  0, 0, &pstm);
        } else {
            hr = pstg->lpVtbl->OpenStream(pstg, Tcl_GetUnicode(objv[2]),
		NULL, mode & STGM_WIN32MASK, 0, &pstm);
            if (FAILED(hr) && mode & STGM_APPEND) {
                hr = pstg->lpVtbl->CreateStream(pstg, Tcl_GetUnicode(objv[2]),
		    mode & STGM_WIN32MASK,  0, 0, &pstm);
            }
        }
	
        if (FAILED(hr)) {
            Tcl_Obj *errObj = Tcl_NewStringObj("", 0);
            Tcl_AppendStringsToObj(errObj, "error opening \"", 
		Tcl_GetString(objv[2]), "\"", (char *)NULL);
	    Tcl_AppendObjToObj(errObj, Win32Error("", hr));
            Tcl_SetObjResult(interp, errObj);
            r = TCL_ERROR;
        } else {
	    Package *pkgPtr;
            StorageChannel *inst;
            char name[3 + TCL_INTEGER_SPACE];
	    
            _snprintf(name, 3 + TCL_INTEGER_SPACE, "stm%ld", 
		InterlockedIncrement(&UNIQUEID));
            inst = (StorageChannel *)ckalloc(sizeof(StorageChannel));
            inst->pstm = pstm;
            inst->grfMode = mode;
            inst->interp = interp;
            inst->watchmask = 0;
	    inst->flags = 0;
	    /* bit0 set then not readable */
            inst->validmask = (mode & STGM_WRITE) ? 0 : TCL_READABLE;
            inst->validmask |= (mode & (STGM_WRITE|STGM_READWRITE)) 
		? TCL_WRITABLE : 0;
            inst->chan = Tcl_CreateChannel(&StorageChannelType, name, 
		    inst, inst->validmask);
            Tcl_RegisterChannel(interp, inst->chan);
            if (mode & STGM_APPEND) {
                Tcl_Seek(inst->chan, 0, SEEK_END);
	    }

	    /* insert at head of channels list */
	    pkgPtr = Tcl_GetAssocData(interp, STORAGE_PACKAGE_KEY, NULL);
	    inst->pkgPtr = pkgPtr;
	    inst->nextPtr = pkgPtr->headPtr;
	    pkgPtr->headPtr = inst;
	    ++pkgPtr->count;

            Tcl_SetObjResult(interp, Tcl_NewStringObj(name, -1));
            r = TCL_OK;
        }
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageStatCmd -
 *
 *	Fetch information about the named item as per [file stat]
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	The array variable passed in will have a number of values added
 *	or modified.
 *
 * ----------------------------------------------------------------------
 */

static int
StorageStatCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    IStorage *pstg = storagePtr->pstg;
    int r = TCL_OK;
    
    if (objc != 4) {

        Tcl_WrongNumArgs(interp, 2, objv, "name varName");
        r = TCL_ERROR;
	
    } else {
	
        STATSTG stat;
        int posixmode = 0;
        const stgm_map_t *p = NULL;
	
        if (r == TCL_OK) {
            r = GetItemInfo(interp, pstg, objv[2], &stat);
            if (r == TCL_OK) {
                Tcl_ObjSetVar2(interp, objv[3], Tcl_NewStringObj("type", -1),
                    (stat.type == STGTY_STORAGE) 
		    ? Tcl_NewStringObj("directory", -1) 
		    : Tcl_NewStringObj("file", -1),
                    0);
                Tcl_ObjSetVar2(interp, objv[3], Tcl_NewStringObj("size", -1),
		    Tcl_NewWideIntObj(stat.cbSize.QuadPart), 0);
                Tcl_ObjSetVar2(interp, objv[3], Tcl_NewStringObj("atime", -1),
		    Tcl_NewLongObj(TimeFromFileTime(&stat.atime)), 0);
                Tcl_ObjSetVar2(interp, objv[3], Tcl_NewStringObj("mtime", -1),
		    Tcl_NewLongObj(TimeFromFileTime(&stat.mtime)), 0);
                Tcl_ObjSetVar2(interp, objv[3], Tcl_NewStringObj("ctime", -1), 
		    Tcl_NewLongObj(TimeFromFileTime(&stat.ctime)), 0);
                Tcl_ObjSetVar2(interp, objv[3], 
		    Tcl_NewStringObj("gid", -1), Tcl_NewLongObj(0), 0);
                Tcl_ObjSetVar2(interp, objv[3], 
		    Tcl_NewStringObj("uid", -1), Tcl_NewLongObj(0), 0);
                Tcl_ObjSetVar2(interp, objv[3], 
		    Tcl_NewStringObj("ino", -1), Tcl_NewLongObj(0), 0);
                Tcl_ObjSetVar2(interp, objv[3], 
		    Tcl_NewStringObj("dev", -1), Tcl_NewLongObj(0), 0);
		
                for (p = stgm_map; p->s != NULL; p++) {
                    if ((storagePtr->mode & ~(STGM_STREAMMASK)) == p->f) {
                        posixmode = p->posixmode;
                        break;
                    }
                }
                Tcl_ObjSetVar2(interp, objv[3], Tcl_NewStringObj("mode", -1),
		    Tcl_NewLongObj(posixmode), 0);
		
                if (stat.pwcsName) {
                    CoTaskMemFree(stat.pwcsName);
                }
            }
        }
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageNamesCmd -
 *
 *	Obtain a list of all item names contained in this storage.
 *
 * Results:
 *	A standard Tcl result. The list of names is returned in the 
 *	interpreters result.
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

static int
StorageNamesCmd(ClientData clientData, Tcl_Interp *interp, 
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    IStorage *pstg = storagePtr->pstg;
    IEnumSTATSTG *penum = NULL;
    STATSTG stats[12];
    ULONG count, n, found = 0;
    int r = TCL_OK;
    
    if (objc > 2) {
	
        Tcl_WrongNumArgs(interp, 2, objv, "");
        r = TCL_ERROR;
	
    } else {
	
        HRESULT hr = pstg->lpVtbl->EnumElements(pstg, 0, NULL, 0, &penum);
        if (FAILED(hr)) {
            Tcl_SetObjResult(interp, Win32Error("names error", hr));
            r = TCL_ERROR;
        } else {
            Tcl_Obj *listObj = Tcl_NewListObj(0, NULL);
            while (hr == S_OK) {
                hr = penum->lpVtbl->Next(penum, 12, stats, &count);
                for (n = 0; SUCCEEDED(hr) && n < count; n++) {
                    Tcl_ListObjAppendElement(interp, listObj, 
                        Tcl_NewUnicodeObj(stats[n].pwcsName, -1));
                    CoTaskMemFree(stats[n].pwcsName);
                }
            }
            penum->lpVtbl->Release(penum);
            Tcl_SetObjResult(interp, listObj);
        }
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageRenameCmd -
 *
 *	Change the name of a storage item.
 *
 * Results:
 *	A standard Tcl result.
 *
 * Side effects:
 *	The items name is changed or an error is raised.
 *
 * ----------------------------------------------------------------------
 */

static int
StorageRenameCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    IStorage *pstg = storagePtr->pstg;
    int r = TCL_OK;
    
    if (objc != 4) {
        Tcl_WrongNumArgs(interp, 2, objv, "oldname newname");
        r = TCL_ERROR;
    } else {
        HRESULT hr = pstg->lpVtbl->RenameElement(pstg, 
            Tcl_GetUnicode(objv[2]), Tcl_GetUnicode(objv[3]));
        if (FAILED(hr)) {
            Tcl_Obj *errObj = Tcl_NewStringObj("", 0);
            Tcl_AppendStringsToObj(errObj, "error renaming \"", 
                Tcl_GetString(objv[2]), 
                "\": no such file or directory", (char *)NULL);
            Tcl_SetObjResult(interp, errObj);
            r = TCL_ERROR;
        }
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageRemoveCmd -
 *
 *	Removes the item from the storage. If the named item is a 
 *	sub-storage then it is removed EVEN IF NOT EMPTY.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	The named item may be deleted from the storage.
 *
 * ----------------------------------------------------------------------
 */

static int
StorageRemoveCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    Storage *storagePtr = (Storage *)clientData;
    IStorage *pstg = storagePtr->pstg;
    int r = TCL_OK;
    
    if (objc != 3) {
        
        Tcl_WrongNumArgs(interp, 2, objv, "name");
        r = TCL_ERROR;
        
    } else {
        
        HRESULT hr = pstg->lpVtbl->DestroyElement(pstg, 
            Tcl_GetUnicode(objv[2]));
        if (FAILED(hr) && hr != STG_E_FILENOTFOUND) {
            Tcl_Obj *errObj = Tcl_NewStringObj("", 0);
            Tcl_AppendStringsToObj(errObj, "error removing \"", 
                Tcl_GetString(objv[2]), "\"", (char *)NULL);
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
 * StorageChannelClose -
 *
 *	Called by the Tcl channel layer to close the channel.
 *	The channel must be removed from the linked list help
 *	in the Package structure.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	Closes the stream and releases allocated resources.
 *
 * ----------------------------------------------------------------------
 */

static int 
StorageChannelClose(ClientData instanceData, Tcl_Interp *interp)
{
    StorageChannel *instPtr = instanceData;
    StorageChannel **tmpPtrPtr;
    Package *pkgPtr = instPtr->pkgPtr;
    
    /* remove this channel from the package list */
    tmpPtrPtr = &pkgPtr->headPtr;
    while (*tmpPtrPtr && *tmpPtrPtr != instPtr) {
	tmpPtrPtr = &(*tmpPtrPtr)->nextPtr;
    }
    *tmpPtrPtr = instPtr->nextPtr;
    --pkgPtr->count;

    /* free the stream and the memory */
    if (instPtr->pstm)
        instPtr->pstm->lpVtbl->Release(instPtr->pstm);
    ckfree((char *)instPtr);
    
    return TCL_OK;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageChannelInput -
 *
 *	Called by the Tcl channel layer to read data from the channel
 *
 * Results:
 *	The number of bytes read or -1 on error.
 *
 * Side effects:
 *	Copies bytes from the stream into the buffer.
 *
 * ----------------------------------------------------------------------
 */

static int 
StorageChannelInput(ClientData instanceData,
    char *buffer, int toRead, int *errorCodePtr)
{
    StorageChannel *chan = (StorageChannel *)instanceData;
    int cb = 0;
    
    if (chan->pstm) {
        HRESULT hr = chan->pstm->lpVtbl->Read(chan->pstm, buffer, toRead, &cb);
        if (FAILED(hr)) {
            cb = -1;
            *errorCodePtr = EINVAL;
        }
    }
    
    return cb;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageChannelOutput -
 *
 *	Called by the Tcl channel layer to write data to the channel.
 *
 * Results:
 *	The number of bytes written or -1 on error.
 *
 * Side effects:
 *	Copies bytes from the buffer into the stream.
 *
 * ----------------------------------------------------------------------
 */

static int 
StorageChannelOutput(ClientData instanceData, 
    CONST84 char *buffer, int toWrite, int *errorCodePtr)
{
    StorageChannel *chan = (StorageChannel *)instanceData;
    int cb = 0;
    
    if (chan->pstm) {
        HRESULT hr = chan->pstm->lpVtbl->Write(chan->pstm, buffer, 
            toWrite, &cb);
        if (FAILED(hr)) {
            cb = -1;
            *errorCodePtr = EINVAL;
        }
    }
    
    return cb;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageChannelSeek -
 *
 *	Called by the Tcl channel layer to change the stream position.
 *
 * Results:
 *	The new seek position.
 *
 * Side effects:
 *	Moves the stream position.
 *
 * ----------------------------------------------------------------------
 */

static int 
StorageChannelSeek(ClientData instanceData,
    long offset, int mode, int *errorCodePtr)
{
    return Tcl_WideAsLong(StorageChannelWideSeek(instanceData, 
        Tcl_LongAsWide(offset), mode, errorCodePtr));
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageChannelWideSeek -
 *
 *	Wide version of the seek operation.
 *
 * Results:
 *	The new seek position as a wide value.
 *
 * Side effects:
 *	Moves the seek position.
 *
 * ----------------------------------------------------------------------
 */

static Tcl_WideInt
StorageChannelWideSeek(ClientData instanceData, Tcl_WideInt offset, 
    int seekMode, int *errorCodePtr)
{
    StorageChannel *chan = (StorageChannel *)instanceData;
    HRESULT hr = S_OK;
    int cb = 0;
    LARGE_INTEGER li; 
    ULARGE_INTEGER uli;
    
    li.QuadPart = offset;
    uli.QuadPart = 0;
    if (chan->pstm) {
        DWORD grfMode = STREAM_SEEK_SET;
        if (seekMode == SEEK_END) 
            grfMode = STREAM_SEEK_END;
        else if (seekMode == SEEK_CUR)
            grfMode = STREAM_SEEK_CUR;
        hr = chan->pstm->lpVtbl->Seek(chan->pstm, li, grfMode, &uli);
        if (FAILED(hr)) {
            *errorCodePtr = EINVAL;
        } else {
            *errorCodePtr = (int)hr;
        }
    }
    return uli.QuadPart;
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageChannelWatch -
 *
 *	Called by the Tcl channel layer when someone calls 'fileevent' on
 *	our channel handle. As we are always readable and writable, we
 *	set the watchmask flag appropriately and set the blocktime to 0
 *	This allows the notified to call SetupProc and CheckProc to
 *	poll any of the channels from this package for events.
 *
 * Results:
 *	None.
 *
 * Side effects:
 *	Will cause the notifier to poll.
 *
 * ----------------------------------------------------------------------
 */

static void
StorageChannelWatch(ClientData instanceData, int mask)
{
    StorageChannel *chan = (StorageChannel *)instanceData;
    Tcl_Time blockTime = { 0, 0 };
    
    /* Set the block time to zero - we are always ready for events. */
    chan->watchmask = mask & chan->validmask;
    if (chan->watchmask) {
        Tcl_SetMaxBlockTime(&blockTime);
    }
}

/*
 * ----------------------------------------------------------------------
 *
 * StorageChannelGetHandle -
 *
 *	Provides a properly COM AddRef'd interface pointer to the 
 *	underlying IStream. The caller is responsible for Release'ing 
 *	this (normal COM rules).
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	An extra reference to the stream is returned to the called.
 *
 * NOTE: If anyone really intends to use this it might be better to
 * 	add the interface pointer to the global interface table and
 *	return the cookie to the caller. This would ensure correct
 *	interface marshalling.
 *
 * ----------------------------------------------------------------------
 */

static int
StorageChannelGetHandle(ClientData instanceData, 
    int direction, ClientData *handlePtr)
{
    StorageChannel *chan = (StorageChannel *)instanceData;
    HRESULT hr = chan->pstm->lpVtbl->QueryInterface(chan->pstm, 
        &IID_IStream, handlePtr);
    return SUCCEEDED(hr) ? TCL_OK : TCL_ERROR;
}

static int
EventProc(Tcl_Event *evPtr, int flags)
{
    ChannelEvent *eventPtr = (ChannelEvent *)evPtr;
    StorageChannel *chanPtr = eventPtr->instPtr;

    if (!(flags & TCL_FILE_EVENTS)) {
	return 0;
    }
    chanPtr->flags &= ~STORAGE_FLAG_PENDING;
    Tcl_NotifyChannel(chanPtr->chan, chanPtr->watchmask & eventPtr->flags);
    return 1;
}

/**
 * This function is called to setup the notifier to monitor our
 * channel for file events. Our CheckProc will be called anyway after some
 * interval so we really only need to ensure that it is called at some 
 * appropriate interval.
 */

static void
SetupProc(ClientData clientData, int flags)
{
    Package *pkgPtr = clientData;
    StorageChannel *chanPtr = NULL;
    int msec = 10000;
    Tcl_Time blockTime = {0, 0};
    
    if (!(flags & TCL_FILE_EVENTS)) {
	return;
    }
    
    for (chanPtr = pkgPtr->headPtr; chanPtr != NULL; chanPtr = chanPtr->nextPtr) {
	msec = 10;
    }
    blockTime.sec = msec / 1000;
    blockTime.usec = (msec % 1000) * 1000;
    Tcl_SetMaxBlockTime(&blockTime);
}

static void
CheckProc(ClientData clientData, int flags)
{
    Package *pkgPtr = clientData;
    StorageChannel *chanPtr = NULL;
    int mask;

    if (!(flags & TCL_FILE_EVENTS)) {
	return;
    }

    for (chanPtr = pkgPtr->headPtr; chanPtr != NULL; chanPtr = chanPtr->nextPtr) {
	if (chanPtr->watchmask == 0) {
	    continue;
	}

	/* queue an event to trigger the notifier - we use an event
	 * for this to avoid starving other resources
	 * We are always writable and readable.
	 */
	mask = TCL_WRITABLE | TCL_READABLE;
	if (chanPtr->watchmask & mask) {
	    ChannelEvent *evPtr = (ChannelEvent *)ckalloc(sizeof(ChannelEvent));
	    chanPtr->flags |= STORAGE_FLAG_PENDING;
	    evPtr->header.proc = EventProc;
	    evPtr->instPtr = chanPtr;
	    evPtr->flags = mask;
	    Tcl_QueueEvent((Tcl_Event *)evPtr, TCL_QUEUE_TAIL);
	}
    }
}

/*
 * ----------------------------------------------------------------------
 *
 * TclEnsembleCmd -
 *
 *	A general purpose ensemble command implementation. This
 *	lets us define a command in terms of it's sub-commands as
 *	a structure.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	A sub-command will be called - anything may happen.
 *
 * ----------------------------------------------------------------------
 */

int
TclEnsembleCmd(ClientData clientData, Tcl_Interp *interp,
    int objc, Tcl_Obj *const objv[])
{
    EnsembleCmdData *data = (EnsembleCmdData *)clientData;
    Ensemble *ensemble = data->ensemble;
    int option = 1;
    int index;
    while (option < objc) {
        if (Tcl_GetIndexFromObjStruct(interp, objv[option], ensemble, 
            sizeof(ensemble[0]), "command", 0, &index) 
            != TCL_OK) 
        {
            return TCL_ERROR;
        }
        if (ensemble[index].command) {
            return ensemble[index].command(data->clientData, 
                interp, objc, objv);
        }
        ensemble = ensemble[index].ensemble;
        option++;
    }
    Tcl_WrongNumArgs(interp, option, objv, "option ?arg arg ...?");
    return TCL_ERROR;
}

/*
 * ----------------------------------------------------------------------
 *
 * GetItemInfo -
 *
 *	Iterate over the items in the storage and return the
 *	STATSTG structure for the matching item or generate
 *	a suitable Tcl error message.
 *
 * Results:
 *	A standard Tcl result
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

static int
GetItemInfo(Tcl_Interp *interp, IStorage *pstg, 
    Tcl_Obj *pathObj, STATSTG *pstatstg)
{
    IEnumSTATSTG *penum = NULL;
    STATSTG stats[12];
    ULONG count, n, objc, found = 0, r = TCL_OK;
    Tcl_Obj **objv;
    HRESULT hr = S_OK;
    
    r = Tcl_ListObjGetElements(interp, pathObj, &objc, &objv);
    if (r == TCL_OK) {
        if (objc < 1) {
            hr = pstg->lpVtbl->Stat(pstg, pstatstg, STATFLAG_DEFAULT);
            found = 1;
        } else {
            LPCOLESTR pwcsName = Tcl_GetUnicode(objv[objc-1]);
            hr = pstg->lpVtbl->EnumElements(pstg, 0, NULL, 0, &penum);
            while (hr == S_OK) {
                hr = penum->lpVtbl->Next(penum, 12, stats, &count);
                for (n = 0; SUCCEEDED(hr) && n < count; n++) {
                    if (!found && wcscmp(pwcsName, stats[n].pwcsName) == 0) {
                        /* we must finish the loop to cleanup the strings */
                        found = 1; 
                        CopyMemory(pstatstg, &stats[n], sizeof(STATSTG));
                        hr = S_FALSE; /* avoid any additional calls to Next */
                    } else {
                        CoTaskMemFree(stats[n].pwcsName);
                    }
                }
            }
        }
        if (penum)
            penum->lpVtbl->Release(penum);
        if (!found) {
            Tcl_SetObjResult(interp, 
                Tcl_NewStringObj("file does not exist", -1));
            r = TCL_ERROR;
        }
    }
    return r;
}

/*
 * ----------------------------------------------------------------------
 *
 * Win32Error -
 *
 *	Convert COM errors into Tcl string objects.
 *
 * Results:
 *	A tcl string object
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

Tcl_Obj *
Win32Error(const char * szPrefix, HRESULT hr)
{
    Tcl_Obj *msgObj = NULL;
    char * lpBuffer = NULL;
    DWORD  dwLen = 0;
    
    /* deal with a few known values */
    switch (hr) {
	case STG_E_FILENOTFOUND: {
	    msgObj = Tcl_NewStringObj(szPrefix, -1);
	    Tcl_AppendToObj(msgObj, ": file not found", -1);
	    return msgObj;
	}
	case STG_E_ACCESSDENIED: {
	    msgObj = Tcl_NewStringObj(szPrefix, -1);
	    Tcl_AppendToObj(msgObj, ": permission denied", -1);
	    return msgObj;
	}
    }

    dwLen = FormatMessageA(FORMAT_MESSAGE_ALLOCATE_BUFFER 
        | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL, (DWORD)hr, LANG_NEUTRAL,
        (LPTSTR)&lpBuffer, 0, NULL);
    if (dwLen < 1) {
        dwLen = FormatMessageA(FORMAT_MESSAGE_ALLOCATE_BUFFER 
            | FORMAT_MESSAGE_FROM_STRING
            | FORMAT_MESSAGE_ARGUMENT_ARRAY,
            "code 0x%1!08X!%n", 0, LANG_NEUTRAL,
            (LPTSTR)&lpBuffer, 0, (va_list *)&hr);
    }
    
    msgObj = Tcl_NewStringObj(szPrefix, -1);
    if (dwLen > 0) {
	char *p = lpBuffer + dwLen - 1;        /* remove cr-lf at end */
	for ( ; p && *p && isspace(*p); p--)
	    ;
	*++p = 0;
	Tcl_AppendToObj(msgObj, ": ", 2);
	Tcl_AppendToObj(msgObj, lpBuffer, -1);
    }
    LocalFree((HLOCAL)lpBuffer);
    return msgObj;
}

/*
 * ----------------------------------------------------------------------
 *
 * TimeToFileTime -
 *
 *	Convert a time_t value into Win32 FILETIME.
 *
 * Results:
 *	The filetime value is modified.
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

static void
TimeToFileTime(time_t t, LPFILETIME pft)
{
    LONGLONG t64 = Int32x32To64(t, 10000000) + 116444736000000000;
    pft->dwLowDateTime = (DWORD)(t64);
    pft->dwHighDateTime = (DWORD)(t64 >> 32);
}

/*
 * ----------------------------------------------------------------------
 *
 * TimeFromFileTime
 *
 *	Convert a FILETIME value into a localtime time_t value.
 *
 * Results:
 *	The localtime in unix epoch seconds.
 *
 * Side effects:
 *	None.
 *
 * ----------------------------------------------------------------------
 */

static time_t
TimeFromFileTime(const FILETIME *pft)
{
    LONGLONG t64 = pft->dwHighDateTime;
    t64 <<= 32;
    t64 |= pft->dwLowDateTime;
    t64 -= 116444736000000000;
    return (time_t)(t64 / 10000000);
}

/* ----------------------------------------------------------------------
 *
 * Local variables:
 * mode: c
 * indent-tabs-mode: nil
 * End:
 */
