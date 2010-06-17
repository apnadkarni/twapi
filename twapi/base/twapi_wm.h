/* Windows messages defined by TWAPI */

/* Messages used by hidden window implementations */

/* Note we now base off of WM_USER, not WM_APP since the latter is meant
   for application-defined messages and the former for control and window
   classes. We, as a library, are more of window class than the application
*/
#define TWAPI_WM_BASE WM_USER
#define TWAPI_WM_HIDDEN_WINDOW_INIT (TWAPI_WM_BASE+0)
#define TWAPI_WM_ADD_HOTKEY         (TWAPI_WM_BASE+1)
#define TWAPI_WM_REMOVE_HOTKEY      (TWAPI_WM_BASE+2)
#define TWAPI_WM_ADD_DIR_NOTIFICATION            (TWAPI_WM_BASE+3)
#define TWAPI_WM_REMOVE_DIR_NOTIFICATION            (TWAPI_WM_BASE+4)
#define TWAPI_WM_ADD_DEVICE_NOTIFICATION    (TWAPI_WM_BASE+5)
#define TWAPI_WM_REMOVE_DEVICE_NOTIFICATION (TWAPI_WM_BASE+6)
