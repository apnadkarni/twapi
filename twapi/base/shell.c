/* 
 * Copyright (c) 2004-2009 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to process information */

#include "twapi.h"


static DLLVERSIONINFO *TwapiShellVersion()
{
    static DLLVERSIONINFO shellver;
    static int initialized = 0;
    
    if (! initialized) {
        TwapiGetDllVersion("shell32.dll", &shellver);
        initialized = 1;
    }
    return &shellver;
}

int Twapi_GetShellVersion(Tcl_Interp *interp)
{
    char buf[80];

    DLLVERSIONINFO *ver = TwapiShellVersion();
    wsprintfA(buf, "%u.%u.%u",
              ver->dwMajorVersion, ver->dwMinorVersion, ver->dwBuildNumber);
    Tcl_SetResult(interp, buf, TCL_VOLATILE);
    return TCL_OK;
}


/* Even though ShGetFolderPath exists on Win2K or later, the VC 6 does
   not like the format of the SDK shell32.lib so we have to stick
   with the VC6 shell32, which does not have this function. So we are
   forced to dynamically load it.
*/
typedef HRESULT (WINAPI *SHGetFolderPathW_t)(HWND, int, HANDLE, DWORD, LPWSTR);
MAKE_DYNLOAD_FUNC(SHGetFolderPathW, shell32, SHGetFolderPathW_t)
HRESULT Twapi_SHGetFolderPath(
    HWND hwndOwner,
    int nFolder,
    HANDLE hToken,
    DWORD flags,
    WCHAR *pathbuf              /* Must be MAX_PATH */
)
{
    SHGetFolderPathW_t SHGetFolderPathPtr = Twapi_GetProc_SHGetFolderPathW();

    if (SHGetFolderPathPtr == NULL) {
        SetLastError(ERROR_PROC_NOT_FOUND);
        return ERROR_PROC_NOT_FOUND;
    }

    return (*SHGetFolderPathPtr)(hwndOwner, nFolder, hToken, flags, pathbuf);
}

typedef BOOL (WINAPI *SHObjectProperties_t)(HWND, DWORD, PCWSTR, PCWSTR);
MAKE_DYNLOAD_FUNC(SHObjectProperties, shell32, SHObjectProperties_t)
MAKE_DYNLOAD_FUNC_ORDINAL(178, shell32)
BOOL Twapi_SHObjectProperties(
    HWND hwnd,
    DWORD dwType,
    LPCWSTR szObject,
    LPCWSTR szPage
)
{
    static SHObjectProperties_t fnSHObjectProperties;
    static int initialized = 0;

    if (! initialized) {
        fnSHObjectProperties = Twapi_GetProc_SHObjectProperties();
        if (fnSHObjectProperties == NULL) {
            /*
             * Could not get function by name. On Win 2K, function is
             * available but not by name. Try getting by ordinal
             * after making sure shell version is 5.0
             */
            DLLVERSIONINFO *ver = TwapiShellVersion();
            if (ver->dwMajorVersion == 5 && ver->dwMinorVersion == 0) {
                fnSHObjectProperties = (SHObjectProperties_t) Twapi_GetProc_shell32_178();
            }
        }
        initialized = 1;
    }

    if (fnSHObjectProperties == NULL) {
        SetLastError(ERROR_PROC_NOT_FOUND);
        return FALSE;
    }

    return (*fnSHObjectProperties)(hwnd, dwType, szObject, szPage);
}

int TwapiGetThemeDefine(Tcl_Interp *interp, char *name)
{
    int val;

    if (name[0] == 0) {
        val = 0;
        goto success_return;
    }

#define cmp_and_return_(def) \
    do { \
        if ( (# def)[0] == name[0] && \
             (lstrcmpi(# def, name) == 0) ) { \
            val = def; \
            goto success_return; \
        } \
    } while (0)

    // TBD - make more efficient by using Tcl hash tables
    // NOTE: some values have been commented out as they are not in my header
    // files though they are listed in MSDN

    cmp_and_return_(BP_CHECKBOX);
    cmp_and_return_(CBS_CHECKEDDISABLED);
    cmp_and_return_(CBS_CHECKEDHOT);
    cmp_and_return_(CBS_CHECKEDNORMAL);
    cmp_and_return_(CBS_CHECKEDPRESSED);
    cmp_and_return_(CBS_MIXEDDISABLED);
    cmp_and_return_(CBS_MIXEDHOT);
    cmp_and_return_(CBS_MIXEDNORMAL);
    cmp_and_return_(CBS_MIXEDPRESSED);
    cmp_and_return_(CBS_UNCHECKEDDISABLED);
    cmp_and_return_(CBS_UNCHECKEDHOT);
    cmp_and_return_(CBS_UNCHECKEDNORMAL);
    cmp_and_return_(CBS_UNCHECKEDPRESSED);
    cmp_and_return_(BP_GROUPBOX);
    cmp_and_return_(GBS_DISABLED);
    cmp_and_return_(GBS_NORMAL);
    cmp_and_return_(BP_PUSHBUTTON);
    cmp_and_return_(PBS_DEFAULTED);
    cmp_and_return_(PBS_DISABLED);
    cmp_and_return_(PBS_HOT);
    cmp_and_return_(PBS_NORMAL);
    cmp_and_return_(PBS_PRESSED);
    cmp_and_return_(BP_RADIOBUTTON);
    cmp_and_return_(RBS_CHECKEDDISABLED);
    cmp_and_return_(RBS_CHECKEDHOT);
    cmp_and_return_(RBS_CHECKEDNORMAL);
    cmp_and_return_(RBS_CHECKEDPRESSED);
    cmp_and_return_(RBS_UNCHECKEDDISABLED);
    cmp_and_return_(RBS_UNCHECKEDHOT);
    cmp_and_return_(RBS_UNCHECKEDNORMAL);
    cmp_and_return_(RBS_UNCHECKEDPRESSED);
    cmp_and_return_(BP_USERBUTTON);
    cmp_and_return_(CLP_TIME);
    cmp_and_return_(CLS_NORMAL);
    cmp_and_return_(CP_DROPDOWNBUTTON);
    cmp_and_return_(CBXS_DISABLED);
    cmp_and_return_(CBXS_HOT);
    cmp_and_return_(CBXS_NORMAL);
    cmp_and_return_(CBXS_PRESSED);
    cmp_and_return_(EP_CARET);
    cmp_and_return_(EP_EDITTEXT);
    cmp_and_return_(ETS_ASSIST);
    cmp_and_return_(ETS_DISABLED);
    cmp_and_return_(ETS_FOCUSED);
    cmp_and_return_(ETS_HOT);
    cmp_and_return_(ETS_NORMAL);
    cmp_and_return_(ETS_READONLY);
    cmp_and_return_(ETS_SELECTED);
    cmp_and_return_(EBP_HEADERBACKGROUND);
    cmp_and_return_(EBP_HEADERCLOSE);
    cmp_and_return_(EBHC_HOT);
    cmp_and_return_(EBHC_NORMAL);
    cmp_and_return_(EBHC_PRESSED);
    cmp_and_return_(EBP_HEADERPIN);
    cmp_and_return_(EBHP_HOT);
    cmp_and_return_(EBHP_NORMAL);
    cmp_and_return_(EBHP_PRESSED);
    cmp_and_return_(EBHP_SELECTEDHOT);
    cmp_and_return_(EBHP_SELECTEDNORMAL);
    cmp_and_return_(EBHP_SELECTEDPRESSED);
    cmp_and_return_(EBP_IEBARMENU);
    cmp_and_return_(EBM_HOT);
    cmp_and_return_(EBM_NORMAL);
    cmp_and_return_(EBM_PRESSED);
    cmp_and_return_(EBP_NORMALGROUPBACKGROUND);
    cmp_and_return_(EBP_NORMALGROUPCOLLAPSE);
    cmp_and_return_(EBNGC_HOT);
    cmp_and_return_(EBNGC_NORMAL);
    cmp_and_return_(EBNGC_PRESSED);
    cmp_and_return_(EBP_NORMALGROUPEXPAND);
    cmp_and_return_(EBNGE_HOT);
    cmp_and_return_(EBNGE_NORMAL);
    cmp_and_return_(EBNGE_PRESSED);
    cmp_and_return_(EBP_NORMALGROUPHEAD);
    cmp_and_return_(EBP_SPECIALGROUPBACKGROUND);
    cmp_and_return_(EBP_SPECIALGROUPCOLLAPSE);
    cmp_and_return_(EBSGC_HOT);
    cmp_and_return_(EBSGC_NORMAL);
    cmp_and_return_(EBSGC_PRESSED);
    cmp_and_return_(EBP_SPECIALGROUPEXPAND);
    cmp_and_return_(EBSGE_HOT);
    cmp_and_return_(EBSGE_NORMAL);
    cmp_and_return_(EBSGE_PRESSED);
    cmp_and_return_(EBP_SPECIALGROUPHEAD);
    //cmp_and_return_(GP_BORDER);
    //cmp_and_return_(BSS_FLAT);
    //cmp_and_return_(BSS_RAISED);
    //cmp_and_return_(BSS_SUNKEN);
    //cmp_and_return_(GP_LINEHORZ);
    //cmp_and_return_(LHS_FLAT);
    //cmp_and_return_(LHS_RAISED);
    //cmp_and_return_(LHS_SUNKEN);
    //cmp_and_return_(GP_LINEVERT);
    //cmp_and_return_(LVS_FLAT);
    //cmp_and_return_(LVS_RAISED);
    //cmp_and_return_(LVS_SUNKEN);
    cmp_and_return_(HP_HEADERITEM);
    cmp_and_return_(HIS_HOT);
    cmp_and_return_(HIS_NORMAL);
    cmp_and_return_(HIS_PRESSED);
    cmp_and_return_(HP_HEADERITEMLEFT);
    cmp_and_return_(HILS_HOT);
    cmp_and_return_(HILS_NORMAL);
    cmp_and_return_(HILS_PRESSED);
    cmp_and_return_(HP_HEADERITEMRIGHT);
    cmp_and_return_(HIRS_HOT);
    cmp_and_return_(HIRS_NORMAL);
    cmp_and_return_(HIRS_PRESSED);
    cmp_and_return_(HP_HEADERSORTARROW);
    cmp_and_return_(HSAS_SORTEDDOWN);
    cmp_and_return_(HSAS_SORTEDUP);
    cmp_and_return_(LVP_EMPTYTEXT);
    cmp_and_return_(LVP_LISTDETAIL);
    cmp_and_return_(LVP_LISTGROUP);
    cmp_and_return_(LVP_LISTITEM);
    cmp_and_return_(LIS_DISABLED);
    cmp_and_return_(LIS_HOT);
    cmp_and_return_(LIS_NORMAL);
    cmp_and_return_(LIS_SELECTED);
    cmp_and_return_(LIS_SELECTEDNOTFOCUS);
    cmp_and_return_(LVP_LISTSORTEDDETAIL);
    cmp_and_return_(MP_MENUBARDROPDOWN);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MP_MENUBARITEM);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MP_CHEVRON);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MP_MENUDROPDOWN);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MP_MENUITEM);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MP_SEPARATOR);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MDP_NEWAPPBUTTON);
    cmp_and_return_(MDS_CHECKED);
    cmp_and_return_(MDS_DISABLED);
    cmp_and_return_(MDS_HOT);
    cmp_and_return_(MDS_HOTCHECKED);
    cmp_and_return_(MDS_NORMAL);
    cmp_and_return_(MDS_PRESSED);
    cmp_and_return_(MDP_SEPERATOR);
    cmp_and_return_(PGRP_DOWN);
    cmp_and_return_(DNS_DISABLED);
    cmp_and_return_(DNS_HOT);
    cmp_and_return_(DNS_NORMAL);
    cmp_and_return_(DNS_PRESSED);
    cmp_and_return_(PGRP_DOWNHORZ);
    cmp_and_return_(DNHZS_DISABLED);
    cmp_and_return_(DNHZS_HOT);
    cmp_and_return_(DNHZS_NORMAL);
    cmp_and_return_(DNHZS_PRESSED);
    cmp_and_return_(PGRP_UP);
    cmp_and_return_(UPS_DISABLED);
    cmp_and_return_(UPS_HOT);
    cmp_and_return_(UPS_NORMAL);
    cmp_and_return_(UPS_PRESSED);
    cmp_and_return_(PGRP_UPHORZ);
    cmp_and_return_(UPHZS_DISABLED);
    cmp_and_return_(UPHZS_HOT);
    cmp_and_return_(UPHZS_NORMAL);
    cmp_and_return_(UPHZS_PRESSED);
    cmp_and_return_(PP_BAR);
    cmp_and_return_(PP_BARVERT);
    cmp_and_return_(PP_CHUNK);
    cmp_and_return_(PP_CHUNKVERT);
    cmp_and_return_(RP_BAND);
    cmp_and_return_(RP_CHEVRON);
    cmp_and_return_(CHEVS_HOT);
    cmp_and_return_(CHEVS_NORMAL);
    cmp_and_return_(CHEVS_PRESSED);
    cmp_and_return_(RP_CHEVRONVERT);
    cmp_and_return_(RP_GRIPPER);
    cmp_and_return_(RP_GRIPPERVERT);
    cmp_and_return_(SBP_ARROWBTN);
    cmp_and_return_(ABS_DOWNDISABLED);
    cmp_and_return_(ABS_DOWNHOT);
    cmp_and_return_(ABS_DOWNNORMAL);
    cmp_and_return_(ABS_DOWNPRESSED);
    cmp_and_return_(ABS_UPDISABLED);
    cmp_and_return_(ABS_UPHOT);
    cmp_and_return_(ABS_UPNORMAL);
    cmp_and_return_(ABS_UPPRESSED);
    cmp_and_return_(ABS_LEFTDISABLED);
    cmp_and_return_(ABS_LEFTHOT);
    cmp_and_return_(ABS_LEFTNORMAL);
    cmp_and_return_(ABS_LEFTPRESSED);
    cmp_and_return_(ABS_RIGHTDISABLED);
    cmp_and_return_(ABS_RIGHTHOT);
    cmp_and_return_(ABS_RIGHTNORMAL);
    cmp_and_return_(ABS_RIGHTPRESSED);
    cmp_and_return_(SBP_GRIPPERHORZ);
    cmp_and_return_(SBP_GRIPPERVERT);
    cmp_and_return_(SBP_LOWERTRACKHORZ);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_LOWERTRACKVERT);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_THUMBBTNHORZ);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_THUMBBTNVERT);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_UPPERTRACKHORZ);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_UPPERTRACKVERT);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_SIZEBOX);
    cmp_and_return_(SZB_LEFTALIGN);
    cmp_and_return_(SZB_RIGHTALIGN);
    cmp_and_return_(SPNP_DOWN);
    cmp_and_return_(DNS_DISABLED);
    cmp_and_return_(DNS_HOT);
    cmp_and_return_(DNS_NORMAL);
    cmp_and_return_(DNS_PRESSED);
    cmp_and_return_(SPNP_DOWNHORZ);
    cmp_and_return_(DNHZS_DISABLED);
    cmp_and_return_(DNHZS_HOT);
    cmp_and_return_(DNHZS_NORMAL);
    cmp_and_return_(DNHZS_PRESSED);
    cmp_and_return_(SPNP_UP);
    cmp_and_return_(UPS_DISABLED);
    cmp_and_return_(UPS_HOT);
    cmp_and_return_(UPS_NORMAL);
    cmp_and_return_(UPS_PRESSED);
    cmp_and_return_(SPNP_UPHORZ);
    cmp_and_return_(UPHZS_DISABLED);
    cmp_and_return_(UPHZS_HOT);
    cmp_and_return_(UPHZS_NORMAL);
    cmp_and_return_(UPHZS_PRESSED);
    cmp_and_return_(SPP_LOGOFF);
    cmp_and_return_(SPP_LOGOFFBUTTONS);
    cmp_and_return_(SPLS_HOT);
    cmp_and_return_(SPLS_NORMAL);
    cmp_and_return_(SPLS_PRESSED);
    cmp_and_return_(SPP_MOREPROGRAMS);
    cmp_and_return_(SPP_MOREPROGRAMSARROW);
    cmp_and_return_(SPS_HOT);
    cmp_and_return_(SPS_NORMAL);
    cmp_and_return_(SPS_PRESSED);
    cmp_and_return_(SPP_PLACESLIST);
    cmp_and_return_(SPP_PLACESLISTSEPARATOR);
    cmp_and_return_(SPP_PREVIEW);
    cmp_and_return_(SPP_PROGLIST);
    cmp_and_return_(SPP_PROGLISTSEPARATOR);
    cmp_and_return_(SPP_USERPANE);
    cmp_and_return_(SPP_USERPICTURE);
    cmp_and_return_(SP_GRIPPER);
    cmp_and_return_(SP_PANE);
    cmp_and_return_(SP_GRIPPERPANE);
    cmp_and_return_(TABP_BODY);
    cmp_and_return_(TABP_PANE);
    cmp_and_return_(TABP_TABITEM);
    cmp_and_return_(TIS_DISABLED);
    cmp_and_return_(TIS_FOCUSED);
    cmp_and_return_(TIS_HOT);
    cmp_and_return_(TIS_NORMAL);
    cmp_and_return_(TIS_SELECTED);
    cmp_and_return_(TABP_TABITEMBOTHEDGE);
    cmp_and_return_(TIBES_DISABLED);
    cmp_and_return_(TIBES_FOCUSED);
    cmp_and_return_(TIBES_HOT);
    cmp_and_return_(TIBES_NORMAL);
    cmp_and_return_(TIBES_SELECTED);
    cmp_and_return_(TABP_TABITEMLEFTEDGE);
    cmp_and_return_(TILES_DISABLED);
    cmp_and_return_(TILES_FOCUSED);
    cmp_and_return_(TILES_HOT);
    cmp_and_return_(TILES_NORMAL);
    cmp_and_return_(TILES_SELECTED);
    cmp_and_return_(TABP_TABITEMRIGHTEDGE);
    cmp_and_return_(TIRES_DISABLED);
    cmp_and_return_(TIRES_FOCUSED);
    cmp_and_return_(TIRES_HOT);
    cmp_and_return_(TIRES_NORMAL);
    cmp_and_return_(TIRES_SELECTED);
    cmp_and_return_(TABP_TOPTABITEM);
    cmp_and_return_(TTIS_DISABLED);
    cmp_and_return_(TTIS_FOCUSED);
    cmp_and_return_(TTIS_HOT);
    cmp_and_return_(TTIS_NORMAL);
    cmp_and_return_(TTIS_SELECTED);
    cmp_and_return_(TABP_TOPTABITEMBOTHEDGE);
    cmp_and_return_(TTIBES_DISABLED);
    cmp_and_return_(TTIBES_FOCUSED);
    cmp_and_return_(TTIBES_HOT);
    cmp_and_return_(TTIBES_NORMAL);
    cmp_and_return_(TTIBES_SELECTED);
    cmp_and_return_(TABP_TOPTABITEMLEFTEDGE);
    cmp_and_return_(TTILES_DISABLED);
    cmp_and_return_(TTILES_FOCUSED);
    cmp_and_return_(TTILES_HOT);
    cmp_and_return_(TTILES_NORMAL);
    cmp_and_return_(TTILES_SELECTED);
    cmp_and_return_(TABP_TOPTABITEMRIGHTEDGE);
    cmp_and_return_(TTIRES_DISABLED);
    cmp_and_return_(TTIRES_FOCUSED);
    cmp_and_return_(TTIRES_HOT);
    cmp_and_return_(TTIRES_NORMAL);
    cmp_and_return_(TTIRES_SELECTED);
    cmp_and_return_(TDP_GROUPCOUNT);
    cmp_and_return_(TDP_FLASHBUTTON);
    cmp_and_return_(TDP_FLASHBUTTONGROUPMENU);
    cmp_and_return_(TBP_BACKGROUNDBOTTOM);
    cmp_and_return_(TBP_BACKGROUNDLEFT);
    cmp_and_return_(TBP_BACKGROUNDRIGHT);
    cmp_and_return_(TBP_BACKGROUNDTOP);
    cmp_and_return_(TBP_SIZINGBARBOTTOM);
    //cmp_and_return_(TBP_SIZINGBARBOTTOMLEFT);
    cmp_and_return_(TBP_SIZINGBARRIGHT);
    cmp_and_return_(TBP_SIZINGBARTOP);
    cmp_and_return_(TP_BUTTON);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TP_DROPDOWNBUTTON);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TP_SPLITBUTTON);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TP_SPLITBUTTONDROPDOWN);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TP_SEPARATOR);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TP_SEPARATORVERT);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TTP_BALLOON);
    cmp_and_return_(TTBS_LINK);
    cmp_and_return_(TTBS_NORMAL);
    cmp_and_return_(TTP_BALLOONTITLE);
    cmp_and_return_(TTBS_LINK);
    cmp_and_return_(TTBS_NORMAL);
    cmp_and_return_(TTP_CLOSE);
    cmp_and_return_(TTCS_HOT);
    cmp_and_return_(TTCS_NORMAL);
    cmp_and_return_(TTCS_PRESSED);
    cmp_and_return_(TTP_STANDARD);
    cmp_and_return_(TTSS_LINK);
    cmp_and_return_(TTSS_NORMAL);
    cmp_and_return_(TTP_STANDARDTITLE);
    cmp_and_return_(TTSS_LINK);
    cmp_and_return_(TTSS_NORMAL);
    cmp_and_return_(TKP_THUMB);
    cmp_and_return_(TUS_DISABLED);
    cmp_and_return_(TUS_FOCUSED);
    cmp_and_return_(TUS_HOT);
    cmp_and_return_(TUS_NORMAL);
    cmp_and_return_(TUS_PRESSED);
    cmp_and_return_(TKP_THUMBBOTTOM);
    cmp_and_return_(TUBS_DISABLED);
    cmp_and_return_(TUBS_FOCUSED);
    cmp_and_return_(TUBS_HOT);
    cmp_and_return_(TUBS_NORMAL);
    cmp_and_return_(TUBS_PRESSED);
    cmp_and_return_(TKP_THUMBLEFT);
    cmp_and_return_(TUVLS_DISABLED);
    cmp_and_return_(TUVLS_FOCUSED);
    cmp_and_return_(TUVLS_HOT);
    cmp_and_return_(TUVLS_NORMAL);
    cmp_and_return_(TUVLS_PRESSED);
    cmp_and_return_(TKP_THUMBRIGHT);
    cmp_and_return_(TUVRS_DISABLED);
    cmp_and_return_(TUVRS_FOCUSED);
    cmp_and_return_(TUVRS_HOT);
    cmp_and_return_(TUVRS_NORMAL);
    cmp_and_return_(TUVRS_PRESSED);
    cmp_and_return_(TKP_THUMBTOP);
    cmp_and_return_(TUTS_DISABLED);
    cmp_and_return_(TUTS_FOCUSED);
    cmp_and_return_(TUTS_HOT);
    cmp_and_return_(TUTS_NORMAL);
    cmp_and_return_(TUTS_PRESSED);
    cmp_and_return_(TKP_THUMBVERT);
    cmp_and_return_(TUVS_DISABLED);
    cmp_and_return_(TUVS_FOCUSED);
    cmp_and_return_(TUVS_HOT);
    cmp_and_return_(TUVS_NORMAL);
    cmp_and_return_(TUVS_PRESSED);
    cmp_and_return_(TKP_TICS);
    cmp_and_return_(TSS_NORMAL);
    cmp_and_return_(TKP_TICSVERT);
    cmp_and_return_(TSVS_NORMAL);
    cmp_and_return_(TKP_TRACK);
    cmp_and_return_(TRS_NORMAL);
    cmp_and_return_(TKP_TRACKVERT);
    cmp_and_return_(TRVS_NORMAL);
    cmp_and_return_(TNP_ANIMBACKGROUND);
    cmp_and_return_(TNP_BACKGROUND);
    cmp_and_return_(TVP_BRANCH);
    cmp_and_return_(TVP_GLYPH);
    cmp_and_return_(GLPS_CLOSED);
    cmp_and_return_(GLPS_OPENED);
    cmp_and_return_(TVP_TREEITEM);
    cmp_and_return_(TREIS_DISABLED);
    cmp_and_return_(TREIS_HOT);
    cmp_and_return_(TREIS_NORMAL);
    cmp_and_return_(TREIS_SELECTED);
    cmp_and_return_(TREIS_SELECTEDNOTFOCUS);
    cmp_and_return_(WP_CAPTION);
    cmp_and_return_(CS_ACTIVE);
    cmp_and_return_(CS_DISABLED);
    cmp_and_return_(CS_INACTIVE);
    cmp_and_return_(WP_CAPTIONSIZINGTEMPLATE);
    cmp_and_return_(WP_CLOSEBUTTON);
    cmp_and_return_(CBS_DISABLED);
    cmp_and_return_(CBS_HOT);
    cmp_and_return_(CBS_NORMAL);
    cmp_and_return_(CBS_PUSHED);
    cmp_and_return_(WP_DIALOG);
    cmp_and_return_(WP_FRAMEBOTTOM);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_FRAMEBOTTOMSIZINGTEMPLATE);
    cmp_and_return_(WP_FRAMELEFT);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_FRAMELEFTSIZINGTEMPLATE);
    cmp_and_return_(WP_FRAMERIGHT);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_FRAMERIGHTSIZINGTEMPLATE);
    cmp_and_return_(WP_HELPBUTTON);
    cmp_and_return_(HBS_DISABLED);
    cmp_and_return_(HBS_HOT);
    cmp_and_return_(HBS_NORMAL);
    cmp_and_return_(HBS_PUSHED);
    cmp_and_return_(WP_HORZSCROLL);
    cmp_and_return_(HSS_DISABLED);
    cmp_and_return_(HSS_HOT);
    cmp_and_return_(HSS_NORMAL);
    cmp_and_return_(HSS_PUSHED);
    cmp_and_return_(WP_HORZTHUMB);
    cmp_and_return_(HTS_DISABLED);
    cmp_and_return_(HTS_HOT);
    cmp_and_return_(HTS_NORMAL);
    cmp_and_return_(HTS_PUSHED);
    //cmp_and_return_(WP_MAX_BUTTON);
    cmp_and_return_(MAXBS_DISABLED);
    cmp_and_return_(MAXBS_HOT);
    cmp_and_return_(MAXBS_NORMAL);
    cmp_and_return_(MAXBS_PUSHED);
    cmp_and_return_(WP_MAXCAPTION);
    cmp_and_return_(MXCS_ACTIVE);
    cmp_and_return_(MXCS_DISABLED);
    cmp_and_return_(MXCS_INACTIVE);
    cmp_and_return_(WP_MDICLOSEBUTTON);
    cmp_and_return_(CBS_DISABLED);
    cmp_and_return_(CBS_HOT);
    cmp_and_return_(CBS_NORMAL);
    cmp_and_return_(CBS_PUSHED);
    cmp_and_return_(WP_MDIHELPBUTTON);
    cmp_and_return_(HBS_DISABLED);
    cmp_and_return_(HBS_HOT);
    cmp_and_return_(HBS_NORMAL);
    cmp_and_return_(HBS_PUSHED);
    cmp_and_return_(WP_MDIMINBUTTON);
    cmp_and_return_(MINBS_DISABLED);
    cmp_and_return_(MINBS_HOT);
    cmp_and_return_(MINBS_NORMAL);
    cmp_and_return_(MINBS_PUSHED);
    cmp_and_return_(WP_MDIRESTOREBUTTON);
    cmp_and_return_(RBS_DISABLED);
    cmp_and_return_(RBS_HOT);
    cmp_and_return_(RBS_NORMAL);
    cmp_and_return_(RBS_PUSHED);
    cmp_and_return_(WP_MDISYSBUTTON);
    cmp_and_return_(SBS_DISABLED);
    cmp_and_return_(SBS_HOT);
    cmp_and_return_(SBS_NORMAL);
    cmp_and_return_(SBS_PUSHED);
    cmp_and_return_(WP_MINBUTTON);
    cmp_and_return_(MINBS_DISABLED);
    cmp_and_return_(MINBS_HOT);
    cmp_and_return_(MINBS_NORMAL);
    cmp_and_return_(MINBS_PUSHED);
    cmp_and_return_(WP_MINCAPTION);
    cmp_and_return_(MNCS_ACTIVE);
    cmp_and_return_(MNCS_DISABLED);
    cmp_and_return_(MNCS_INACTIVE);
    cmp_and_return_(WP_RESTOREBUTTON);
    cmp_and_return_(RBS_DISABLED);
    cmp_and_return_(RBS_HOT);
    cmp_and_return_(RBS_NORMAL);
    cmp_and_return_(RBS_PUSHED);
    cmp_and_return_(WP_SMALLCAPTION);
    cmp_and_return_(CS_ACTIVE);
    cmp_and_return_(CS_DISABLED);
    cmp_and_return_(CS_INACTIVE);
    cmp_and_return_(WP_SMALLCAPTIONSIZINGTEMPLATE);
    cmp_and_return_(WP_SMALLCLOSEBUTTON);
    cmp_and_return_(CBS_DISABLED);
    cmp_and_return_(CBS_HOT);
    cmp_and_return_(CBS_NORMAL);
    cmp_and_return_(CBS_PUSHED);
    cmp_and_return_(WP_SMALLFRAMEBOTTOM);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_SMALLFRAMEBOTTOMSIZINGTEMPLATE);
    cmp_and_return_(WP_SMALLFRAMELEFT);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_SMALLFRAMELEFTSIZINGTEMPLATE);
    cmp_and_return_(WP_SMALLFRAMERIGHT);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_SMALLFRAMERIGHTSIZINGTEMPLATE);
    //cmp_and_return_(WP_SMALLHELPBUTTON);
    cmp_and_return_(HBS_DISABLED);
    cmp_and_return_(HBS_HOT);
    cmp_and_return_(HBS_NORMAL);
    cmp_and_return_(HBS_PUSHED);
    //cmp_and_return_(WP_SMALLMAXBUTTON);
    cmp_and_return_(MAXBS_DISABLED);
    cmp_and_return_(MAXBS_HOT);
    cmp_and_return_(MAXBS_NORMAL);
    cmp_and_return_(MAXBS_PUSHED);
    cmp_and_return_(WP_SMALLMAXCAPTION);
    cmp_and_return_(MXCS_ACTIVE);
    cmp_and_return_(MXCS_DISABLED);
    cmp_and_return_(MXCS_INACTIVE);
    cmp_and_return_(WP_SMALLMINCAPTION);
    cmp_and_return_(MNCS_ACTIVE);
    cmp_and_return_(MNCS_DISABLED);
    cmp_and_return_(MNCS_INACTIVE);
    //cmp_and_return_(WP_SMALLRESTOREBUTTON);
    cmp_and_return_(RBS_DISABLED);
    cmp_and_return_(RBS_HOT);
    cmp_and_return_(RBS_NORMAL);
    cmp_and_return_(RBS_PUSHED);
    //cmp_and_return_(WP_SMALLSYSBUTTON);
    cmp_and_return_(SBS_DISABLED);
    cmp_and_return_(SBS_HOT);
    cmp_and_return_(SBS_NORMAL);
    cmp_and_return_(SBS_PUSHED);
    cmp_and_return_(WP_SYSBUTTON);
    cmp_and_return_(SBS_DISABLED);
    cmp_and_return_(SBS_HOT);
    cmp_and_return_(SBS_NORMAL);
    cmp_and_return_(SBS_PUSHED);
    cmp_and_return_(WP_VERTSCROLL);
    cmp_and_return_(VSS_DISABLED);
    cmp_and_return_(VSS_HOT);
    cmp_and_return_(VSS_NORMAL);
    cmp_and_return_(VSS_PUSHED);
    cmp_and_return_(WP_VERTTHUMB);
    cmp_and_return_(VTS_DISABLED);
    cmp_and_return_(VTS_HOT);
    cmp_and_return_(VTS_NORMAL);
    cmp_and_return_(VTS_PUSHED); 
    
    // GetThemeColor

    cmp_and_return_(TMT_BORDERCOLOR);
    cmp_and_return_(TMT_FILLCOLOR);
    cmp_and_return_(TMT_TEXTCOLOR);
    cmp_and_return_(TMT_EDGELIGHTCOLOR);
    cmp_and_return_(TMT_EDGEHIGHLIGHTCOLOR);
    cmp_and_return_(TMT_EDGESHADOWCOLOR);
    cmp_and_return_(TMT_EDGEDKSHADOWCOLOR);
    cmp_and_return_(TMT_EDGEFILLCOLOR);
    cmp_and_return_(TMT_TRANSPARENTCOLOR);
    cmp_and_return_(TMT_GRADIENTCOLOR1);
    cmp_and_return_(TMT_GRADIENTCOLOR2);
    cmp_and_return_(TMT_GRADIENTCOLOR3);
    cmp_and_return_(TMT_GRADIENTCOLOR4);
    cmp_and_return_(TMT_GRADIENTCOLOR5);
    cmp_and_return_(TMT_SHADOWCOLOR);
    cmp_and_return_(TMT_GLOWCOLOR);
    cmp_and_return_(TMT_TEXTBORDERCOLOR);
    cmp_and_return_(TMT_TEXTSHADOWCOLOR);
    cmp_and_return_(TMT_GLYPHTEXTCOLOR);
    cmp_and_return_(TMT_GLYPHTRANSPARENTCOLOR);
    cmp_and_return_(TMT_FILLCOLORHINT);
    cmp_and_return_(TMT_BORDERCOLORHINT);
    cmp_and_return_(TMT_ACCENTCOLORHINT);
    cmp_and_return_(TMT_BLENDCOLOR);

    // GetThemeFont
    cmp_and_return_(TMT_GLYPHTYPE);
    cmp_and_return_(TMT_FONT);

    Tcl_SetObjErrorCode(interp,
                        Twapi_MakeTwapiErrorCodeObj(TWAPI_INVALID_ARGS));
    Tcl_AppendResult(interp, "Invalid theme symbol '", name, "'", NULL);
    return TCL_ERROR;

 success_return:
    Tcl_SetObjResult(interp, Tcl_NewIntObj(val));
    return TCL_OK;

#undef cmp_and_return_
}



// Create a shell link
int Twapi_WriteShortcut (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR linkPath;
    LPCWSTR objPath;
    LPITEMIDLIST itemIds = NULL;
    LPCWSTR commandArgs;
    LPCWSTR desc;
    WORD    hotkey;
    LPCWSTR iconPath;
    int     iconIndex;
    LPCWSTR relativePath;
    int     showCommand;
    LPCWSTR workingDirectory;

    HRESULT hres; 
    IShellLinkW* psl = NULL; 
    IPersistFile* ppf = NULL;
 
    if (TwapiGetArgs(interp, objc, objv,
                     GETWSTR(linkPath), GETNULLIFEMPTY(objPath),
                     GETVAR(itemIds, ObjToPIDL),
                     GETNULLIFEMPTY(commandArgs),
                     GETNULLIFEMPTY(desc),
                     GETWORD(hotkey),
                     GETNULLIFEMPTY(iconPath),
                     GETINT(iconIndex),
                     GETNULLIFEMPTY(relativePath),
                     GETINT(showCommand),
                     GETNULLIFEMPTY(workingDirectory),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (objPath == NULL && itemIds == NULL)
        return ERROR_INVALID_PARAMETER;

    // Get a pointer to the IShellLink interface. 
    hres = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, 
                            &IID_IShellLinkW, (LPVOID*)&psl); 
    if (FAILED(hres)) 
        return hres;

    if (objPath)
        hres = psl->lpVtbl->SetPath(psl,objPath); 
    if (FAILED(hres))
        goto vamoose;
    if (itemIds)
        hres = psl->lpVtbl->SetIDList(psl, itemIds);
    if (FAILED(hres))
        goto vamoose;
    if (commandArgs)
        hres = psl->lpVtbl->SetArguments(psl, commandArgs);
    if (FAILED(hres))
        goto vamoose;
    if (desc)
        hres = psl->lpVtbl->SetDescription(psl, desc); 
    if (FAILED(hres))
        goto vamoose;
    if (hotkey)
        hres = psl->lpVtbl->SetHotkey(psl, hotkey);
    if (FAILED(hres))
        goto vamoose;
    if (iconPath)
        hres = psl->lpVtbl->SetIconLocation(psl, iconPath, iconIndex);
    if (FAILED(hres))
        goto vamoose;
    if (relativePath)
        hres = psl->lpVtbl->SetRelativePath(psl, relativePath, 0);
    if (FAILED(hres))
        goto vamoose;
    if (showCommand >= 0)
        hres = psl->lpVtbl->SetShowCmd(psl, showCommand);
    if (FAILED(hres))
        goto vamoose;
    if (workingDirectory)
        hres = psl->lpVtbl->SetWorkingDirectory(psl, workingDirectory);
    if (FAILED(hres))
        goto vamoose;

    hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf); 
    if (FAILED(hres))
        goto vamoose;
 
    /* Save the link  */
    hres = ppf->lpVtbl->Save(ppf, linkPath, TRUE); 
    ppf->lpVtbl->Release(ppf); 
    
 vamoose:
    if (psl)
        psl->lpVtbl->Release(psl); 
    TwapiFreePIDL(itemIds);     /* OK if NULL */

    if (hres != S_OK) {
        Twapi_AppendSystemError(interp, hres);
        return TCL_ERROR;
    } else
        return TCL_OK;
}

/* Read a shortcut */
int Twapi_ReadShortcut(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR linkPath;
    int pathFlags;
    HWND hwnd;
    DWORD resolve_flags;

    HRESULT hres; 
    IShellLinkW *psl = NULL; 
    IPersistFile *ppf = NULL;
    Tcl_Obj *resultObj = NULL;
#if (INFOTIPSIZE > MAX_PATH)
    WCHAR buf[INFOTIPSIZE+1];
#else
    WCHAR buf[MAX_PATH+1];
#endif
    WORD  wordval;
    int   intval;
    LPITEMIDLIST pidl;
    int   retval = TCL_ERROR;
 
    if (TwapiGetArgs(interp, objc, objv,
                     GETWSTR(linkPath), GETINT(pathFlags),
                     GETDWORD_PTR(hwnd), GETINT(resolve_flags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* Get a pointer to the IShellLink interface. */
    hres = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, 
                            &IID_IShellLinkW, (LPVOID*)&psl); 
    if (FAILED(hres))
        goto fail;

    /* Load the resource through the IPersist interface */
    hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf); 
    if (FAILED(hres))
        goto fail;
    
    hres = ppf->lpVtbl->Load(ppf, linkPath, STGM_READ);
    if (FAILED(hres))
        goto fail;

    /* Resolve the link */
    hres = psl->lpVtbl->Resolve(psl, hwnd, resolve_flags);
#if 0    /* Ignore resolve errors */
    if (FAILED(hres))
        goto fail;
#endif

    resultObj = Tcl_NewListObj(0, NULL);

    /*
     * Get each field. Note that inability to get a field is not treated
     * as an error. We just go on and try to get the next one
     */
    hres = psl->lpVtbl->GetArguments(psl, buf, sizeof(buf)/sizeof(buf[0]));
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-args"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewUnicodeObj(buf, -1));
    }

    hres = psl->lpVtbl->GetDescription(psl, buf, sizeof(buf)/sizeof(buf[0]));
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-desc"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewUnicodeObj(buf, -1));
    }

    hres = psl->lpVtbl->GetHotkey(psl, &wordval);
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-hotkey"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewIntObj(wordval));
    }

    hres = psl->lpVtbl->GetIconLocation(psl,
                                        buf, sizeof(buf)/sizeof(buf[0]),
                                        &intval);
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-iconindex"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewIntObj(intval));
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-iconpath"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewUnicodeObj(buf, -1));
    }

    hres = psl->lpVtbl->GetIDList(psl, &pidl);
    if (hres == NOERROR) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-idl"));
        Tcl_ListObjAppendElement(interp, resultObj, ObjFromPIDL(pidl));
        CoTaskMemFree(pidl);
    }

    hres = psl->lpVtbl->GetPath(psl, buf, sizeof(buf)/sizeof(buf[0]),
                                NULL, pathFlags);
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-path"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewUnicodeObj(buf, -1));
    }

    hres = psl->lpVtbl->GetShowCmd(psl, &intval);
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-showcmd"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewIntObj(intval));
    }

    hres = psl->lpVtbl->GetWorkingDirectory(psl, buf, sizeof(buf)/sizeof(buf[0]));
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-workdir"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewUnicodeObj(buf, -1));
    }
    
    Tcl_SetObjResult(interp, resultObj);
    retval = TCL_OK;

 vamoose:
    if (psl)
        psl->lpVtbl->Release(psl); 
    if (ppf)
        ppf->lpVtbl->Release(ppf); 

    return retval;

 fail:
    if (resultObj)
        Twapi_FreeNewTclObj(resultObj);
    resultObj = NULL;
    Twapi_AppendSystemError(interp, hres);
    goto vamoose;
}


// Create a url link
int Twapi_WriteUrlShortcut(Tcl_Interp *interp, LPCWSTR linkPath, LPCWSTR url, DWORD flags)
{ 
    HRESULT hres; 
    IUniformResourceLocatorW *psl = NULL; 
    IPersistFile* ppf = NULL;
 
    // Get a pointer to the IShellLink interface. 
    hres = CoCreateInstance(&CLSID_InternetShortcut, NULL,
                            CLSCTX_INPROC_SERVER, 
                            &IID_IUniformResourceLocatorW, (LPVOID*)&psl); 

    if (FAILED(hres)) {
        /* No interface and hence no interface specific error, just return as standard error */
        return Twapi_AppendSystemError(interp, hres);
    }

    hres = psl->lpVtbl->SetURL(psl, url, flags);
    if (FAILED(hres)) {
        TWAPI_STORE_COM_ERROR(interp, hres, psl, &IID_IUniformResourceLocatorW);
    } else {
        hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf);
        if (FAILED(hres)) {
            /* No-op - this is a standard error so we do not get try
               getting ISupportErrorInfo */
            Twapi_AppendSystemError(interp, hres);
        } else {
            /* Save the link  */
            hres = ppf->lpVtbl->Save(ppf, linkPath, TRUE); 
            if (FAILED(hres)) {
                TWAPI_STORE_COM_ERROR(interp, hres, ppf, &IID_IPersistFile);
            }
        }
    }

    if (ppf)
        ppf->lpVtbl->Release(ppf); 
    if (psl)
        psl->lpVtbl->Release(psl); 

    return FAILED(hres) ? TCL_ERROR : TCL_OK;
}

/* Read a URL shortcut */
int Twapi_ReadUrlShortcut(Tcl_Interp *interp, LPCWSTR linkPath)
{
    HRESULT hres; 
    IUniformResourceLocatorW *psl = NULL; 
    IPersistFile *ppf = NULL;
    LPWSTR url;
    int   retval = TCL_ERROR;
 
    /* Get a pointer to the IShellLink interface. */
    hres = CoCreateInstance(&CLSID_InternetShortcut, NULL,
                            CLSCTX_INPROC_SERVER, 
                            &IID_IUniformResourceLocatorW, (LPVOID*)&psl); 
    if (FAILED(hres))
        goto fail;

    /* Load the resource through the IPersist interface */
    hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf); 
    if (FAILED(hres))
        goto fail;
    
    hres = ppf->lpVtbl->Load(ppf, linkPath, STGM_READ);
    if (FAILED(hres))
        goto fail;

    hres = psl->lpVtbl->GetURL(psl, &url);

    if (FAILED(hres))
        goto fail;

    Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(url, -1));
    CoTaskMemFree(url);

    retval = TCL_OK;

 vamoose:
    if (psl)
        psl->lpVtbl->Release(psl); 
    if (ppf)
        ppf->lpVtbl->Release(ppf); 

    return retval;

 fail:
    Twapi_AppendSystemError(interp, hres);
    goto vamoose;
}


/* Invoke a URL shortcut */
int Twapi_InvokeUrlShortcut(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR linkPath;
    LPCWSTR verb;
    DWORD flags;
    HWND hwnd;

    HRESULT hres; 
    IUniformResourceLocatorW *psl = NULL; 
    IPersistFile *ppf = NULL;
    URLINVOKECOMMANDINFOW urlcmd;
 
    if (TwapiGetArgs(interp, objc, objv,
                     GETWSTR(linkPath), GETWSTR(verb), GETINT(flags),
                     GETHANDLE(hwnd),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* Get a pointer to the IShellLink interface. */
    hres = CoCreateInstance(&CLSID_InternetShortcut, NULL,
                            CLSCTX_INPROC_SERVER, 
                            &IID_IUniformResourceLocatorW, (LPVOID*)&psl); 
    if (SUCCEEDED(hres)) {
        hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf); 
    }
    if (FAILED(hres)) {
        /* This CoCreateInstance or QueryInterface error so we do not get
           use TWAPI_STORE_COM_ERROR */
        Twapi_AppendSystemError(interp, hres);
    } else {
        hres = ppf->lpVtbl->Load(ppf, linkPath, STGM_READ);
        if (FAILED(hres)) {
            TWAPI_STORE_COM_ERROR(interp, hres, ppf, &IID_IPersistFile);
        } else {
            urlcmd.dwcbSize = sizeof(urlcmd);
            urlcmd.dwFlags = flags;
            urlcmd.hwndParent = hwnd;
            urlcmd.pcszVerb = verb;

            hres = psl->lpVtbl->InvokeCommand(psl, &urlcmd);
            if (FAILED(hres)) {
                TWAPI_STORE_COM_ERROR(interp, hres, psl, &IID_IUniformResourceLocatorW);
            }
        }
    }

    if (ppf)
        ppf->lpVtbl->Release(ppf); 
    if (psl)
        psl->lpVtbl->Release(psl); 

    return FAILED(hres) ? TCL_ERROR : TCL_OK;
}


MAKE_DYNLOAD_FUNC(OpenThemeData, uxtheme, FARPROC)
HTHEME Twapi_OpenThemeData(HWND win, LPCWSTR classes)
{
    FARPROC fn;

    fn = Twapi_GetProc_OpenThemeData();
    return fn ? (HTHEME) (*fn)(win, classes) : NULL;
}

MAKE_DYNLOAD_FUNC(CloseThemeData, uxtheme, FARPROC)
void Twapi_CloseThemeData(HTHEME themeH)
{
    FARPROC fn = Twapi_GetProc_CloseThemeData();
    if (fn)
        (void) (*fn)(themeH);
}

MAKE_DYNLOAD_FUNC(IsThemeActive, uxtheme, FARPROC_BOOL)
BOOL Twapi_IsThemeActive(void)
{
    FARPROC_BOOL fn = Twapi_GetProc_IsThemeActive();
    return fn ? (*fn)() : FALSE;
}

MAKE_DYNLOAD_FUNC(IsAppThemed, uxtheme, FARPROC_BOOL)
BOOL Twapi_IsAppThemed(void)
{
    FARPROC_BOOL fn = Twapi_GetProc_IsAppThemed();
    return fn ? (*fn)() : FALSE;
}

typedef HRESULT (WINAPI *GetCurrentThemeName_t)(LPWSTR, int, LPWSTR, int, LPWSTR, int);
MAKE_DYNLOAD_FUNC(GetCurrentThemeName, uxtheme, GetCurrentThemeName_t)
int Twapi_GetCurrentThemeName(Tcl_Interp *interp)
{
    WCHAR filename[MAX_PATH];
    WCHAR color[256];
    WCHAR size[256];
    GetCurrentThemeName_t fn = Twapi_GetProc_GetCurrentThemeName();
    HRESULT status;
    Tcl_Obj *objv[3];

    if (fn == NULL) {
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
    }

    status = (*fn)(filename, ARRAYSIZE(filename),
                   color, ARRAYSIZE(color),
                   size, ARRAYSIZE(size));

    if (status != S_OK) {
        return Twapi_AppendSystemError(interp, status);
    }

    objv[0] = Tcl_NewUnicodeObj(filename, -1);
    objv[1] = Tcl_NewUnicodeObj(color, -1);
    objv[2] = Tcl_NewUnicodeObj(size, -1);
    Tcl_SetObjResult(interp, Tcl_NewListObj(3, objv));
    return TCL_OK;
}

typedef HRESULT (WINAPI *GetThemeColor_t)(HTHEME, int, int, int, COLORREF *);
MAKE_DYNLOAD_FUNC(GetThemeColor, uxtheme, GetThemeColor_t)
int Twapi_GetThemeColor(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HTHEME hTheme;
    int iPartId;
    int iStateId;
    int iPropId;

    GetThemeColor_t fn = Twapi_GetProc_GetThemeColor();
    HRESULT status;
    COLORREF color;
    char buf[40];

    if (fn == NULL)
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(hTheme), GETINT(iPartId), GETINT(iStateId),
                     GETINT(iPropId),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    status =  (*fn)(hTheme, iPartId, iStateId, iPropId, &color);
    if (status != S_OK)
        return Twapi_AppendSystemError(interp, status);

    wsprintfA(buf, "#%2.2x%2.2x%2.2x",
             GetRValue(color), GetGValue(color), GetBValue(color));
    Tcl_SetObjResult(interp, Tcl_NewStringObj(buf, -1));
    return TCL_OK;
}


typedef HRESULT (WINAPI *GetThemeFont_t)(HTHEME, HDC, int, int, int, LOGFONTW *);
MAKE_DYNLOAD_FUNC(GetThemeFont, uxtheme, GetThemeFont_t)
int Twapi_GetThemeFont(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HTHEME hTheme;
    HANDLE hdc;
    int iPartId;
    int iStateId;
    int iPropId;

    LOGFONTW lf;
    HRESULT hr;

    GetThemeFont_t fn = Twapi_GetProc_GetThemeFont();

    /* NOTE GetThemeFont ExPECTS LOGFONTW although the documentation
     * mentions LOGFONT
     */
    if (fn == NULL)
        return ERROR_PROC_NOT_FOUND;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(hTheme),  GETHANDLE(hdc),
                     GETINT(iPartId), GETINT(iStateId), GETINT(iPropId),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    hr =  (*fn)(hTheme, hdc, iPartId, iStateId, iPropId, &lf);
    if (hr != S_OK)
        return Twapi_AppendSystemError(interp, hr);

    Tcl_SetObjResult(interp, ObjFromLOGFONTW(&lf));
    return TCL_OK;
}


int Twapi_SHFileOperation (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    SHFILEOPSTRUCTW sfop;
    Tcl_Obj *objs[2];
    int      tcl_status = TCL_ERROR;

    sfop.pFrom = NULL;          /* To track necessary deallocs */
    sfop.pTo   = NULL;
    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(sfop.hwnd), GETINT(sfop.wFunc),
                     GETVAR(sfop.pFrom, ObjToMultiSz),
                     GETVAR(sfop.pTo, ObjToMultiSz),
                     GETWORD(sfop.fFlags), GETWSTR(sfop.lpszProgressTitle),
                     ARGEND) != TCL_OK)
        goto vamoose;

    sfop.hNameMappings = NULL;

    if (SHFileOperationW(&sfop) != 0) {
        // Note GetLastError() is not set by the call
        Tcl_SetResult(interp, "SHFileOperation failed", TCL_STATIC);
        goto vamoose;
    }

    objs[0] = Tcl_NewBooleanObj(sfop.fAnyOperationsAborted);
    objs[1] = Tcl_NewListObj(0, NULL);

    if (sfop.hNameMappings) {
        int i;
        SHNAMEMAPPINGW *mapP = *(SHNAMEMAPPINGW **)(((char *)sfop.hNameMappings) + 4);
        for (i = 0; i < *(int *) (sfop.hNameMappings); ++i) {
            Tcl_ListObjAppendElement(interp, objs[1],
                                     Tcl_NewUnicodeObj(
                                         mapP[i].pszOldPath,
                                         mapP[i].cchOldPath));
            Tcl_ListObjAppendElement(interp, objs[1],
                                     Tcl_NewUnicodeObj(
                                         mapP[i].pszNewPath,
                                         mapP[i].cchNewPath));
        }

        SHFreeNameMappings(sfop.hNameMappings);
    }

    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objs));
    tcl_status = TCL_OK;

vamoose:
    if (sfop.pFrom)
        TwapiFree((void*)sfop.pFrom);
    if (sfop.pTo)
        TwapiFree((void*)sfop.pTo);

    return tcl_status;
}

int Twapi_ShellExecuteEx(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR lpClass;
    HKEY hkeyClass;
    DWORD dwHotKey;
    HANDLE hIconOrMonitor;

    SHELLEXECUTEINFOW sei;
    
    ZeroMemory(&sei, sizeof(sei)); /* Also sets sei.lpIDList = NULL -
                                     Need to track if it needs freeing */
    sei.cbSize = sizeof(sei);

    if (TwapiGetArgs(interp, objc, objv,
                     GETINT(sei.fMask), GETHANDLE(sei.hwnd),
                     GETNULLIFEMPTY(sei.lpVerb),
                     GETNULLIFEMPTY(sei.lpFile),
                     GETNULLIFEMPTY(sei.lpParameters),
                     GETNULLIFEMPTY(sei.lpDirectory),
                     GETINT(sei.nShow),
                     GETVAR(sei.lpIDList, ObjToPIDL),
                     GETNULLIFEMPTY(lpClass),
                     GETHANDLE(hkeyClass),
                     GETINT(dwHotKey),
                     GETHANDLE(hIconOrMonitor),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (sei.fMask & SEE_MASK_CLASSNAME)
        sei.lpClass = lpClass;
    if (sei.fMask & SEE_MASK_CLASSKEY)
        sei.hkeyClass = hkeyClass;
    if (sei.fMask & SEE_MASK_HOTKEY)
        sei.dwHotKey = dwHotKey;
    if (sei.fMask & SEE_MASK_ICON)
        sei.hIcon = hIconOrMonitor;
    if (sei.fMask & SEE_MASK_HMONITOR)
        sei.hMonitor = hIconOrMonitor;

    if (ShellExecuteExW(&sei) == 0) {
        Tcl_Obj *objP = Tcl_NewStringObj("ShellExecuteEx specific error: ", -1);
        Tcl_AppendObjToObj(objP, ObjFromDWORD_PTR(sei.hInstApp));
        TwapiFreePIDL(sei.lpIDList);     /* OK if NULL */
        return Twapi_AppendSystemError2(interp, GetLastError(), objP);
    }

    TwapiFreePIDL(sei.lpIDList);     /* OK if NULL */

    /* Success, see if any fields to be returned */
    if (sei.fMask & SEE_MASK_NOCLOSEPROCESS) {
        Tcl_SetObjResult(interp, ObjFromHANDLE(sei.hProcess));
    }
    return TCL_OK;
}


int Twapi_SHChangeNotify(
    ClientData notused,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    LONG event_id;
    UINT flags;
    LPVOID dwItem1 = NULL;
    LPVOID dwItem2 = NULL;
    LPITEMIDLIST idl1P = NULL;
    LPITEMIDLIST idl2P = NULL;
    int status = TCL_ERROR;


    if (objc < 3) {
        goto wrong_nargs_error;
    }

    if (Tcl_GetLongFromObj(interp, objv[1], &event_id) != TCL_OK ||
        Tcl_GetIntFromObj(interp, objv[2], &flags) != TCL_OK) {
        goto vamoose;
    }
        
    switch (flags & SHCNF_TYPE) {
    case SHCNF_DWORD:
    case SHCNF_PATHW:
    case SHCNF_PRINTERW:
    case SHCNF_IDLIST:
        /* Valid but no special treatment */
        break;
    case SHCNF_PATHA:
        /* Always pass as Unicode */
        flags = (flags & ~SHCNF_TYPE) | SHCNF_PATHW;
        break;
    case SHCNF_PRINTERA:
        /* Always pass as Unicode */
        flags = (flags & ~SHCNF_TYPE) | SHCNF_PRINTERW;
        break;
        
    default:
        goto invalid_flags_error;
    }


    switch (event_id) {
    case SHCNE_ASSOCCHANGED:
        /* Both dwItem1 and dwItem2 must be NULL (already set) */
        if (! (flags & SHCNF_IDLIST)) {
            /* SDK says this should be set */
            goto invalid_flags_error;
        }
        break;

    case SHCNE_ATTRIBUTES:
    case SHCNE_CREATE:
    case SHCNE_DELETE:
    case SHCNE_DRIVEADD:
    case SHCNE_DRIVEADDGUI:
    case SHCNE_DRIVEREMOVED:
    case SHCNE_FREESPACE:
    case SHCNE_MEDIAINSERTED:
    case SHCNE_MEDIAREMOVED:
    case SHCNE_MKDIR:
    case SHCNE_RMDIR:
    case SHCNE_NETSHARE:
    case SHCNE_NETUNSHARE:
    case SHCNE_UPDATEDIR:
    case SHCNE_UPDATEITEM:
    case SHCNE_SERVERDISCONNECT:
        /* For the above, only dwItem1 used, dwItem2 remains 0 */

        if (objc < 4)
            goto wrong_nargs_error;

        switch (flags & SHCNF_TYPE) {
        case SHCNF_IDLIST:
            if (ObjToPIDL(interp, objv[3], &idl1P) != TCL_OK)
                goto vamoose;
            dwItem1 = idl1P;
            break;
        case SHCNF_PATHW:
            dwItem1 = Tcl_GetUnicode(objv[3]);
            break;
        default:
            goto invalid_flags_error;
        }

        break;

    case SHCNE_RENAMEITEM:
    case SHCNE_RENAMEFOLDER:
        /* Both dwItem1 and dwItem2 used */
        if (objc < 5)
            goto wrong_nargs_error;

        switch (flags & SHCNF_TYPE) {
        case SHCNF_IDLIST:
            if (ObjToPIDL(interp, objv[3], &idl1P) != TCL_OK ||
                ObjToPIDL(interp, objv[4], &idl2P) != TCL_OK)
                goto vamoose;
            dwItem1 = idl1P;
            dwItem2 = idl2P;
            break;
        case SHCNF_PATHW:
            dwItem1 = Tcl_GetUnicode(objv[3]);
            dwItem2 = Tcl_GetUnicode(objv[4]);
            break;
        default:
            goto invalid_flags_error;
        }

        break;

    case SHCNE_UPDATEIMAGE:
        /* dwItem1 not used, dwItem2 is a DWORD */
        if (objc < 5)
            goto wrong_nargs_error;
        if (Tcl_GetLongFromObj(interp, objv[4], (long *)&dwItem2) != TCL_OK)
            goto vamoose;
        break;
        
    case SHCNE_ALLEVENTS:
    case SHCNE_DISKEVENTS:
    case SHCNE_GLOBALEVENTS:
    case SHCNE_INTERRUPT:
        /*
         * SDK docs do not really say what parameters are valid and
         * how to interpret them. So for the below cases we make
         * a best guess based on the parameter and flags supplied by the
         * caller. If number of parameters is less than 4, then assume
         * dwItem1&2 are both unused etc.
         */
        if (objc >= 4) {
            switch (flags & SHCNF_TYPE) {
            case SHCNF_DWORD:
                if (Tcl_GetLongFromObj(interp, objv[3], (long *)&dwItem1) != TCL_OK)
                    goto vamoose;
                if (objc > 4 &&
                    (Tcl_GetLongFromObj(interp, objv[4],
                                        (long *)&dwItem2) != TCL_OK))
                    goto vamoose;
                break;
            case SHCNF_IDLIST:
                if (ObjToPIDL(interp, objv[3], &idl1P) != TCL_OK)
                    goto vamoose;
                dwItem1 = idl1P;
                if (objc > 4) {
                    if (ObjToPIDL(interp, objv[4], &idl2P) != TCL_OK)
                        goto vamoose;
                    dwItem2 = idl2P;
                }
                break;
            case SHCNF_PATHW:
                dwItem1 = Tcl_GetUnicode(objv[3]);
                if (objc > 4)
                    dwItem2 = Tcl_GetUnicode(objv[4]);
                break;
            default:
                goto invalid_flags_error;
            }
        }

        break;

    default:
        Tcl_SetResult(interp, "Unknown or unsupported SHChangeNotify event type", TCL_STATIC);
        goto vamoose;

    }

    /* Note SHChangeNotify has no error return */
    SHChangeNotify(event_id, flags, dwItem1, dwItem2);
    status =  TCL_OK;

vamoose:
    if (idl1P)
        TwapiFreePIDL(idl1P);
    if (idl2P)
        TwapiFreePIDL(idl2P);

    return status;

invalid_flags_error:
    Tcl_SetResult(interp,
                  "Unknown or unsupported SHChangeNotify flags type",
                  TCL_STATIC);
    goto vamoose;

wrong_nargs_error:
    Tcl_WrongNumArgs(interp, 1, objv, "wEventID uFlags dwItem1 ?dwItem2?");
    goto vamoose;
}
