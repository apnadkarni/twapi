#
# Copyright (c) 2007-2010 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Stuff dealing with Microsoft Installer

namespace eval twapi {
    # Holds the MSI prototypes indexed by name
    variable msiprotos_installer
    variable msiprotos_database
    variable msiprotos_record
    variable msi_guids
    array set msi_guids {
        installer {{000C1090-0000-0000-C000-000000000046}}
        database {{000C109D-0000-0000-C000-000000000046}}
        record {{000C1093-0000-0000-C000-000000000046}}
        summaryinfo {{000C109B-0000-0000-C000-000000000046}}
        stringlist {{000C1095-0000-0000-C000-000000000046}}
        view {{000C109C-0000-0000-C000-000000000046}}
            
    }
}

# Initialize MSI module
proc twapi::init_msi {} {

    # Load all the prototypes
    
    # Installer object
    set ::twapi::msiprotos_installer {
        AddSource            {43  0 1 void {bstr bstr bstr}}
        ApplyPatch           {22  0 1 void {bstr bstr i4 bstr}}
        ApplyMultiplePatches {51  0 1 void {bstr bstr bstr}}
        ClearSourceList      {44  0 1 void      {bstr bstr}}
        CollectUserInfo      {21  0 1 void      {bstr}}
        ComponentClients     {38  0 1 idispatch {bstr}}
        ComponentPath        {31  0 1 bstr      {bstr bstr}}
        ComponentQualifiers  {34  0 1 idispatch {bstr}}
        Components           {37  0 1 idispatch {bstr}}
        ConfigureFeature     {28  0 1 void      {bstr bstr bstr}}
        ConfigureProduct     {19  0 1 void      {bstr bstr bstr}}
        CreateRecord         {1   0 1 idispatch {i4}}
        EnableLog            {7   0 1 void      {bstr bstr}}
        Environment          {12  0 2 bstr      {bstr}}
        Environment          {12  0 4 void      {bstr bstr}}
        ExtractPatchXMLData  {57  0 1 void      {bstr}}
        FeatureParent        {23  0 2 bstr      {bstr bstr}}
        Features             {36  0 2 idispatch {bstr}}
        FeatureState         {24  0 2 i4        {bstr bstr}}
        FeatureUsageCount    {26  0 2 i4        {bstr bstr}}
        FeatureUsageDate     {27  0 2 date      {bstr bstr}}
        FileAttributes       {13  0 1 i4        {bstr}}
        FileHash             {47  0 1 idispatch {bstr i4}}
        FileSignatureInfo    {48  0 1 {safearray ui1} {bstr i4 i4}}
        FileSize             {15  0 1 i4       {bstr}}
        FileVersion          {16  0 1 bstr     {bstr {bool {in 0}}}}
        ForceSourceListResolution {45 0 1 void {bstr bstr}}
        InstallProduct       {8   0 1 void      {bstr bstr}}
        LastErrorRecord      {10  0 1 idispatch {}}
        OpenPackage          {2   0 1 idispatch {bstr i4}}
        OpenDatabase         {4   0 1 idispatch {bstr i4}}
        OpenProduct          {3   0 1 idispatch {bstr}}
        PatchInfo            {41  0 2 bstr      {bstr bstr}}
        Patches              {39  0 2 idispatch {bstr}}
        PatchesEx            {55  0 2 idispatch {bstr bstr i4 i4}}
        PatchTransforms      {42  0 2 bstr      {bstr bstr}}
        ProductInfo          {18  0 2 bstr      {bstr bstr}}
        ProductsEx           {52  0 2 idispatch {bstr bstr i4}}
        Products             {35  0 2 idispatch {}}
        ProductState         {17  0 2 bstr      {bstr}}
        ProvideComponent     {30  0 1 bstr      {bstr bstr bstr i4}}
        ProvideQualifiedComponent {32 0 1 bstr  {bstr bstr i4}}
	QualifierDescription {33  0 2 bstr      {bstr bstr}}
        RegistryValue        {11  0 1 bstr      {bstr bstr bstr}}
        ReinstallFeature     {29  0 1 void      {bstr bstr bstr}}
        ReinstallProduct     {20  0 1 void      {bstr bstr}}
        RelatedProducts      {40  0 2 idispatch {bstr}}
        RemovePatches        {49  0 1 void      {bstr bstr i4 bstr}}
        ShortcutTarget       {46  0 2 idispatch {bstr}}
        SummaryInformation   {5   0 2 idispatch {bstr i4}}
        UILevel              {6   0 2 bstr      {}}
        UILevel              {6   0 4 void      {bstr}}
        UseFeature           {25  0 1 void      {bstr bstr bstr}}
        Version              {9   0 2 bstr      {}}
    }

    # Database object
    set ::twapi::msiprotos_database {
        ApplyTransform       {10  0 1 void      {bstr i4}}
        Commit               {4   0 1 void      {}}
        CreateTransformSummaryInfo {13 0 1 void {idispatch bstr i4 i4}}
        DatabaseState        {1   0 2 i4        {}}
        EnableUIPreview      {11  0 1 void      {}}
        Export               {7   0 1 void      {bstr bstr bstr}}
        GenerateTransform    {9   0 1 bool      {idispatch {bstr 0}}}
        Import               {6   0 1 void      {bstr bstr}}
        Merge                {8   0 1 bool      {idispatch {bstr 0}}}
        OpenView             {3   0 1 idispatch  {bstr}}
        PrimaryKeys          {5   0 2 idispatch {bstr}}
        SummaryInformation   {2   0 2 idispatch {i4}}
        TablePersistent      {12  0 2 i4       {bstr}}
    }

    # Record object
    set ::twapi::msiprotos_record {
        ClearData    {7   0 1 void      {}}
        DataSize     {5   0 2 i4        {}}
        FieldCount   {0   0 2 i4        {}}
        FormatText   {8   0 1 void      {}}
        IntegerData  {2   0 2 i4        {i4}}
        IntegerData  {2   0 4 void      {i4 i4}}
        IsNull       {6   0 2 bool      {i4}}
        ReadStream   {4   0 1 bstr      {i4 i4 i4}}
        SetStream    {3   0 1 void      {i4 bstr}}
        StringData   {1   0 2 bstr      {i4}}
        StringData   {1   0 4 void      {i4 bstr}}
    }

    # SummaryInfo
    set ::twapi::msiprotos_summaryinfo {
        Persist    {3   0 1 void      {}}   
        Property   {1   0 2 bstr      {i4}}
        Property   {1   0 4 bstr      {i4}}
        PropertyCount {2   0 2 i4     {}}
    }

    # StringList
    set ::twapi::msiprotos_stringlist {
        Count  {1  0 2 i4   {}}
        Item   {0  0 2 bstr {i4}} 
    }

    # View
    # Flags - 0x11 for Exectute indicate optional,in parameter
    set ::twapi::msiprotos_view {
        Close      {4   0 1 void      {}}
        ColumnInfo {5   0 2 idispatch {i4}}
        Execute    {1   0 1 void      {{9 {0x11}}}}
        Fetch      {2   0 1 idispatch {}}
        GetError   {6   0 1 void      {}}
        Modify     {3   0 1 void      {i4 idispatch}}
    }
}

# Get the MSI installer
proc twapi::new_msi {} {
    return [comobj WindowsInstaller.Installer]
}

# Get the MSI installer
proc twapi::delete_msi {obj} {
    $obj -destroy
}

# Loads msi prototypes for the object
proc twapi::load_msi_prototypes {obj type} {

    # Init protos and stuff
    init_msi

    # Redefine ourselves so we don't call init_msi everytime
    proc ::twapi::load_msi_prototypes {obj type} {
        set type [string tolower $type]
        variable msi_guids
        variable msiprotos_$type

        # Tell the object it's type (guid)
        $obj -interfaceguid $msi_guids($type)

        # Load the prototypes in the cache
        _dispatch_prototype_load $msi_guids($type) [set msiprotos_$type]
    }

    # Call our new definition
    return [load_msi_prototypes $obj $type]
}
