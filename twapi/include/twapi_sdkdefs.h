
/* 
 * Depending on compiler / platform, ZeroMemory translates a call to memset.
 * If we want to avoid the C rtl, then define appropriately
 */
#ifdef TWAPI_REPLACE_CRT
# ifdef _M_AMD64
#  define TwapiZeroMemory(p_, count_) RtlSecureZeroMemory((p_), (count_))
# else
#  define TwapiZeroMemory(p_, count_) ZeroMemory((p_), (count_))
# endif
#else
# define TwapiZeroMemory(p_, count_) memset((p_), 0, (count_))
#endif

/*
 * Contains copies of definitions from newer SDK's.
 * Since we are stuck with the older SDK's for compatibility reasons,
 * we are forced to clone the definitions.
 */
typedef enum _TWAPI_TOKEN_INFORMATION_CLASS {
    TwapiTokenUser = 1,
    TwapiTokenGroups,
    TwapiTokenPrivileges,
    TwapiTokenOwner,
    TwapiTokenPrimaryGroup,
    TwapiTokenDefaultDacl,
    TwapiTokenSource,
    TwapiTokenType,
    TwapiTokenImpersonationLevel,
    TwapiTokenStatistics,
    TwapiTokenRestrictedSids,
    TwapiTokenSessionId,
    TwapiTokenGroupsAndPrivileges,
    TwapiTokenSessionReference,
    TwapiTokenSandBoxInert,
    TwapiTokenAuditPolicy,
    TwapiTokenOrigin,
    TwapiTokenElevationType,
    TwapiTokenLinkedToken,
    TwapiTokenElevation,
    TwapiTokenHasRestrictions,
    TwapiTokenAccessInformation,
    TwapiTokenVirtualizationAllowed,
    TwapiTokenVirtualizationEnabled,
    TwapiTokenIntegrityLevel,
    TwapiTokenUIAccess,
    TwapiTokenMandatoryPolicy,
    TwapiTokenLogonSid,
    MaxTwapiTokenInfoClass  // MaxTokenInfoClass should always be the last enum
} TWAPI_TOKEN_INFORMATION_CLASS;

typedef enum _TWAPI_TOKEN_ELEVATION_TYPE {
    TwapiTokenElevationTypeDefault = 1,
    TwapiTokenElevationTypeFull,
    TwapiTokenElevationTypeLimited,
} TWAPI_TOKEN_ELEVATION_TYPE;

typedef struct _TWAPI_TOKEN_ELEVATION {
    DWORD TokenIsElevated;
} TWAPI_TOKEN_ELEVATION;

typedef struct _TWAPI_TOKEN_LINKED_TOKEN {  
    HANDLE LinkedToken;
} TWAPI_TOKEN_LINKED_TOKEN;

typedef struct _TWAPI_TOKEN_MANDATORY_LABEL {
  SID_AND_ATTRIBUTES Label;
} TWAPI_TOKEN_MANDATORY_LABEL;

typedef struct _TWAPI_TOKEN_MANDATORY_POLICY {
  DWORD Policy;
} TWAPI_TOKEN_MANDATORY_POLICY;

#ifndef SYSTEM_MANDATORY_LABEL_ACE_TYPE
#define SYSTEM_MANDATORY_LABEL_ACE_TYPE (0x11)
#endif
