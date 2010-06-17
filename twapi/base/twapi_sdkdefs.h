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
