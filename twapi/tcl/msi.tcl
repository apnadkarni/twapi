#
# Copyright (c) 2007 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Stuff dealing with Microsoft Installer

namespace eval twapi {
    # Holds the MSI prototypes indexed by name
    variable msiprotos_installer
    variable msiprotos_database
    variable msiprotos_record
}

# Initialize MSI module
proc twapi::init_msi {} {

    # Load all the prototypes
    
    # Installer object
    foreach {name proto} {
        AddSource            {43  0 1 void {bstr bstr bstr}}
        ApplyPatch           TBD
        ApplyMultiplePatches TBD
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
        FileAttributes       {13  0 2 i4        {bstr}}
        FileHash             TBD
        FileSignatureInfo    TBD
        FileSize             {15  0 1 i4       {bstr}}
        FileVersion          {16  0 1 bstr     {bstr {bool {in 0}}}}
        ForceSourceListResolution TBD
        InstallProduct       {8   0 1 void      {bstr bstr}}
        LastErrorRecord      {10  0 1 idispatch {}}
        OpenPackage          {2   0 1 idispatch {bstr i4}}
        OpenDatabase         {4   0 1 idispatch {bstr i4}}
        OpenProduct          {3   0 1 idispatch {bstr}}
        Patches              {39  0 2 idispatch {bstr}}
        PatchesEx            {55  0 2 idispatch {bstr bstr i4 i4}}
        PatchInfo            TBD
        PatchTransforms      TBD
        ProductInfo          {18  0 2 bstr      {bstr bstr}}
        ProductsEx           {52  0 2 idispatch {bstr bstr i4}}
        Products             {35  0 2 idispatch {}}
        ProductState         {17  0 2 bstr      {bstr}}
        ProvideComponent     {30  0 1 bstr      {bstr bstr bstr i4}}
        ProvideQualifiedComponent     TBD
        QualifierDescription TBD
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
    } {
        # We skip TBD's
        if {[llength $proto] > 1} {
            set ::twapi::msiprotos_installer($name) $proto
        }
    }

    # Database object
    foreach {name proto} {
        ApplyTransform       {10  0 1 void      {bstr i4}}
        Commit               {4   0 1 void      {}}
        CreateTransformSummaryInfo TBD-13
        DatabaseState        {1   0 2 i4        {}}
        EnableUIPreview      {11  0 1 void      {}}
        Export               {7   0 1 void      {bstr bstr bstr}}
        GenerateTransform    TBD-9
        Import               {6   0 1 void      {bstr bstr}}
        Merge                TBD-8
        OpenView             {3   0 1 idispatch  {bstr}}
        PrimaryKeys          {5   0 2 idispatch {bstr}}
        SummaryInformation   {2   0 2 idispatch {i4}}
        TablePersistent      {12  0 2 i4       {bstr}}
    } {
        # We skip TBD's
        if {[llength $proto] > 1} {
            set ::twapi::msiprotos_database($name) $proto
        }
    }

    # Record object
    foreach {name proto} {
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
    } {
        # We skip TBD's
        if {[llength $proto] > 1} {
            set ::twapi::msiprotos_record($name) $proto
        }
    }

    # SummaryInfo
    foreach {name proto} {
        Persist    {3   0 1 void      {}}   
        Property   {1   0 2 bstr      {i4}}
        Property   {1   0 4 bstr      {i4}}
        PropertyCount {2   0 2 i4     {}}
    } {
        # We skip TBD's
        if {[llength $proto] > 1} {
            set ::twapi::msiprotos_summaryinfo($name) $proto
        }
    }

    # StringList
    foreach {name proto} {
        Count  {1  0 2 i4   {}}
        Item   {0  0 2 bstr {i4}} 
    } {
        # We skip TBD's
        if {[llength $proto] > 1} {
            set ::twapi::msiprotos_stringlist($name) $proto
        }
    }

    # View
    # Flags - 0x11 for Exectute indicate optional,in parameter
    foreach {name proto} {
        Close      {4   0 1 void      {}}
        ColumnInfo {5   0 2 idispatch {i4}}
        Execute    {1   0 1 void      {{9 {0x11}}}}
        Fetch      {2   0 1 idispatch {}}
        GetError   {6   0 1 void      {}}
        Modify     {3   0 1 void      {i4 idispatch}}
    } {
        # We skip TBD's
        if {[llength $proto] > 1} {
            set ::twapi::msiprotos_view($name) $proto
        }
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
        variable msiprotos_[string tolower $type]

        $obj -precache [array get msiprotos_[string tolower $type]]
    }

    # Call our new definition
    return [load_msi_prototypes $obj $type]
}
