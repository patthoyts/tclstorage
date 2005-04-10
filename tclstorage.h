/* tclstorage.h - Copyright (C) 2005 Pat Thoyts <patthoyts@users.sf.net>
 *
 * $Id$
 */

#define WIN32_LEAN_AND_MEAN
#define STRICT
#include <ole2.h>
#include <tcl.h>
#include <errno.h>
#include <time.h>

#undef TCL_STORAGE_CLASS
#define TCL_STORAGE_CLASS DLLEXPORT

typedef struct Ensemble {
    const char *name;           /* subcommand name */
    Tcl_ObjCmdProc *command;    /* implementation OR */
    struct Ensemble *ensemble;  /* subcommand ensemble */
} Ensemble;

typedef struct EnsembleCmdData {
    struct Ensemble *ensemble;
    ClientData       clientData;
} EnsembleCmdData;

typedef struct {
    IStorage *pstg;
    int       mode;
    Tcl_Obj  *children;
} Storage;

#define STGM_APPEND     0x00000004  /* unused bit in Win32 enum */
#define STGM_TRUNC      0x00004000  /*   "            "         */
#define STGM_WIN32MASK  0xFFFFBFFB  /* mask to remove private bits */
#define STGM_STREAMMASK 0xFFFFAFF8  /* mask off the access, create and
                                       append bits */

EXTERN int Storage_Init(Tcl_Interp *interp);
EXTERN int Storage_SafeInit(Tcl_Interp *interp);
EXTERN Tcl_ObjCmdProc Storage_OpenStorage;

int GetStorageFlagsFromObj(Tcl_Interp *interp, Tcl_Obj *objPtr, int *flagsPtr);
Tcl_ObjCmdProc StoragePropertySetCmd;
Tcl_ObjCmdProc TclEnsembleCmd;
Tcl_Obj *Win32Error(const char * szPrefix, HRESULT hr);
