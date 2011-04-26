/* Windows messages defined by TWAPI */

/* Messages used by hidden window implementations */

/* Note we now base off of WM_USER, not WM_APP since the latter is meant
   for application-defined messages and the former for control and window
   classes. We, as a library, are more of window class than the application
*/
#define TWAPI_WM_BASE WM_USER
#define TWAPI_WM_HIDDEN_WINDOW_INIT (TWAPI_WM_BASE+0)
#define TWAPI_WM_ADD_HOTKEY         (TWAPI_WM_BASE+1) //TBD OBSOLETE?
#define TWAPI_WM_REMOVE_HOTKEY      (TWAPI_WM_BASE+2) //TBD OBSOLETE?
#define TWAPI_WM_ADD_DIR_NOTIFICATION            (TWAPI_WM_BASE+3) //TBD OBSOLETE?
#define TWAPI_WM_REMOVE_DIR_NOTIFICATION            (TWAPI_WM_BASE+4) //TBD OBSOLETE?
#define TWAPI_WM_ADD_DEVICE_NOTIFICATION    (TWAPI_WM_BASE+5)
#define TWAPI_WM_REMOVE_DEVICE_NOTIFICATION (TWAPI_WM_BASE+6)

/*
 * Reserve message id that can be used directly at script level with
 * the default hidden notification window. If you change this, make
 * sure to change the _wm_script_messages variable at script level.
 */
#define TWAPI_WM_SCRIPT_BASE    (TWAPI_WM_BASE + 32)
#define TWAPI_WM_SCRIPT_LAST    (TWAPI_WM_SCRIPT_BASE + 31)

