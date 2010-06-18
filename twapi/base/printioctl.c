/*
 * Program to print ioctl codes and guids.
 * Takes one optional argument and prints the corresponding value.
 * If no arguments, prints all values
 */

#include <windows.h>
#include <initguid.h>
#include <winioctl.h>
#include <stdio.h>
#include <setupapi.h>
#include <objbase.h>
#include <devguid.h>
#include <stddef.h>

/* Macro to print offsets of structures
   To use, modify source to add a line like:
    printstructoffset(_SP_DEVINFO_DATA, ClassGuid);
*/


#define PRINTNL printf("\n")


/* From DDK headers that we do not / cannot include */
DEFINE_GUID(GUID_NDIS_LAN_CLASS, 0xad498944, 0x762f, 0x11d0, 0x8d, 0xcb, 0x00, 0xc0, 0x4f, 0xc3, 0x35, 0x8c); // ndisguid.h

int main(int argc, char **argv)
{
    wchar_t guidstr[40];
    char *wanted = NULL;


    if (argc > 1)
        wanted = argv[1];

#define printioctl(_sym) \
    do { \
        if (wanted == NULL || !stricmp(wanted, "codes") || !stricmp(wanted, # _sym)) \
            printf("%-32s: %d, 0x%x\n", # _sym, _sym, _sym); \
    } while (0)

    printioctl(IOCTL_CHANGER_EXCHANGE_MEDIUM);
    printioctl(IOCTL_CHANGER_GET_ELEMENT_STATUS);
    printioctl(IOCTL_CHANGER_GET_PARAMETERS);
    printioctl(IOCTL_CHANGER_GET_PRODUCT_DATA);
    printioctl(IOCTL_CHANGER_GET_STATUS);
    printioctl(IOCTL_CHANGER_INITIALIZE_ELEMENT_STATUS);
    printioctl(IOCTL_CHANGER_MOVE_MEDIUM);
    printioctl(IOCTL_CHANGER_QUERY_VOLUME_TAGS);
    printioctl(IOCTL_CHANGER_REINITIALIZE_TRANSPORT);
    printioctl(IOCTL_CHANGER_SET_POSITION);



    printioctl(IOCTL_DISK_CHECK_VERIFY);
    printioctl(IOCTL_DISK_CONTROLLER_NUMBER);
    printioctl(IOCTL_DISK_CREATE_DISK);
    printioctl(IOCTL_DISK_DELETE_DRIVE_LAYOUT);
    printioctl(IOCTL_DISK_EJECT_MEDIA);
    printioctl(IOCTL_DISK_FIND_NEW_DEVICES);
    printioctl(IOCTL_DISK_FORMAT_DRIVE);
    printioctl(IOCTL_DISK_FORMAT_TRACKS);
    printioctl(IOCTL_DISK_FORMAT_TRACKS_EX);
    printioctl(IOCTL_DISK_GET_CACHE_INFORMATION);
    printioctl(IOCTL_DISK_GET_DRIVE_GEOMETRY);
    printioctl(IOCTL_DISK_GET_DRIVE_GEOMETRY_EX);
    printioctl(IOCTL_DISK_GET_DRIVE_LAYOUT);
    printioctl(IOCTL_DISK_GET_DRIVE_LAYOUT_EX);
    printioctl(IOCTL_DISK_GET_LENGTH_INFO);
    printioctl(IOCTL_DISK_GET_MEDIA_TYPES);
    printioctl(IOCTL_DISK_GET_PARTITION_INFO);
    printioctl(IOCTL_DISK_GET_PARTITION_INFO_EX);
    printioctl(IOCTL_DISK_GROW_PARTITION);
    printioctl(IOCTL_DISK_IS_WRITABLE);
    printioctl(IOCTL_DISK_HISTOGRAM_DATA);
    printioctl(IOCTL_DISK_HISTOGRAM_RESET);
    printioctl(IOCTL_DISK_HISTOGRAM_STRUCTURE);
    printioctl(IOCTL_DISK_LOGGING);
    printioctl(IOCTL_DISK_PERFORMANCE);
    printioctl(IOCTL_DISK_PERFORMANCE_OFF);
    printioctl(IOCTL_DISK_REASSIGN_BLOCKS);
    printioctl(IOCTL_DISK_REQUEST_DATA);
    printioctl(IOCTL_DISK_REQUEST_STRUCTURE);
    printioctl(IOCTL_DISK_SENSE_DEVICE);
    printioctl(IOCTL_DISK_SET_CACHE_INFORMATION);
    printioctl(IOCTL_DISK_SET_DRIVE_LAYOUT);
    printioctl(IOCTL_DISK_SET_DRIVE_LAYOUT_EX);
    printioctl(IOCTL_DISK_SET_PARTITION_INFO);
    printioctl(IOCTL_DISK_SET_PARTITION_INFO_EX);
    printioctl(IOCTL_DISK_UPDATE_DRIVE_SIZE);
    printioctl(IOCTL_DISK_UPDATE_PROPERTIES);
    printioctl(IOCTL_DISK_VERIFY);

    printioctl(IOCTL_DISK_LOAD_MEDIA);
    printioctl(IOCTL_DISK_MEDIA_REMOVAL);
    printioctl(IOCTL_DISK_RELEASE);
    printioctl(IOCTL_DISK_RESERVE);

    printioctl(IOCTL_SERENUM_EXPOSE_HARDWARE);
    printioctl(IOCTL_SERENUM_GET_PORT_NAME);
    printioctl(IOCTL_SERENUM_PORT_DESC);
    printioctl(IOCTL_SERENUM_REMOVE_HARDWARE);

    printioctl(IOCTL_STORAGE_BREAK_RESERVATION);
    printioctl(IOCTL_STORAGE_CHECK_VERIFY);
    printioctl(IOCTL_STORAGE_CHECK_VERIFY2);
    printioctl(IOCTL_STORAGE_EJECT_MEDIA);
    printioctl(IOCTL_STORAGE_EJECTION_CONTROL);
    printioctl(IOCTL_STORAGE_FIND_NEW_DEVICES);
    printioctl(IOCTL_STORAGE_GET_DEVICE_NUMBER);
    printioctl(IOCTL_STORAGE_GET_HOTPLUG_INFO);
    printioctl(IOCTL_STORAGE_GET_MEDIA_SERIAL_NUMBER);
    printioctl(IOCTL_STORAGE_GET_MEDIA_TYPES);
    printioctl(IOCTL_STORAGE_GET_MEDIA_TYPES_EX);
    printioctl(IOCTL_STORAGE_MEDIA_REMOVAL);
    printioctl(IOCTL_STORAGE_LOAD_MEDIA);
    printioctl(IOCTL_STORAGE_LOAD_MEDIA2);
    printioctl(IOCTL_STORAGE_MCN_CONTROL);
    printioctl(IOCTL_STORAGE_PREDICT_FAILURE);
    printioctl(IOCTL_STORAGE_READ_CAPACITY);
    printioctl(IOCTL_STORAGE_RELEASE);
    printioctl(IOCTL_STORAGE_RESERVE);
    printioctl(IOCTL_STORAGE_RESET_BUS);
    printioctl(IOCTL_STORAGE_RESET_DEVICE);
    printioctl(IOCTL_STORAGE_SET_HOTPLUG_INFO);

    printioctl(IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS);
    printioctl(IOCTL_VOLUME_IS_CLUSTERED);

    printioctl(SMART_GET_VERSION);
    printioctl(SMART_SEND_DRIVE_COMMAND);
    printioctl(SMART_RCV_DRIVE_DATA);

    /* Now do the GUID's */
#define printguid(_sym) \
    do { \
        if (wanted == NULL || !stricmp(wanted,"guids") || !stricmp(wanted, # _sym)) { \
            if (! StringFromGUID2(& _sym, guidstr, sizeof(guidstr)/sizeof(guidstr[0]))) \
                printf("Error converting %s\n", # _sym); \
            else \
                wprintf(L"%-38S: %s\n", # _sym, guidstr);   \
        } \
    } while (0)

    PRINTNL;

    printguid(GUID_DEVINTERFACE_DISK);
    printguid(GUID_DEVINTERFACE_CDROM);
    printguid(GUID_DEVINTERFACE_PARTITION);
    printguid(GUID_DEVINTERFACE_TAPE);
    printguid(GUID_DEVINTERFACE_WRITEONCEDISK);
    printguid(GUID_DEVINTERFACE_VOLUME);
    printguid(GUID_DEVINTERFACE_MEDIUMCHANGER);
    printguid(GUID_DEVINTERFACE_FLOPPY);
    printguid(GUID_DEVINTERFACE_CDCHANGER);
    printguid(GUID_DEVINTERFACE_STORAGEPORT);
    printguid(GUID_DEVINTERFACE_COMPORT);
    printguid(GUID_DEVINTERFACE_SERENUM_BUS_ENUMERATOR);

    printguid(GUID_NDIS_LAN_CLASS);

    /* Device class guids */
    printguid(GUID_DEVCLASS_1394);
    printguid(GUID_DEVCLASS_1394DEBUG);
    printguid(GUID_DEVCLASS_61883);
    printguid(GUID_DEVCLASS_ADAPTER);
    printguid(GUID_DEVCLASS_APMSUPPORT);
    printguid(GUID_DEVCLASS_AVC);
    printguid(GUID_DEVCLASS_BATTERY);
    printguid(GUID_DEVCLASS_BIOMETRIC);
    printguid(GUID_DEVCLASS_BLUETOOTH);
    printguid(GUID_DEVCLASS_CDROM);
    printguid(GUID_DEVCLASS_COMPUTER);
    printguid(GUID_DEVCLASS_DECODER);
    printguid(GUID_DEVCLASS_DISKDRIVE);
    printguid(GUID_DEVCLASS_DISPLAY);
    printguid(GUID_DEVCLASS_DOT4);
    printguid(GUID_DEVCLASS_DOT4PRINT);
    printguid(GUID_DEVCLASS_ENUM1394);
    printguid(GUID_DEVCLASS_FDC);
    printguid(GUID_DEVCLASS_FLOPPYDISK);
    printguid(GUID_DEVCLASS_GPS);
    printguid(GUID_DEVCLASS_HDC);
    printguid(GUID_DEVCLASS_HIDCLASS);
    printguid(GUID_DEVCLASS_IMAGE);
    printguid(GUID_DEVCLASS_INFINIBAND);
    printguid(GUID_DEVCLASS_INFRARED);
    printguid(GUID_DEVCLASS_KEYBOARD);
    printguid(GUID_DEVCLASS_LEGACYDRIVER);
    printguid(GUID_DEVCLASS_MEDIA);
    printguid(GUID_DEVCLASS_MEDIUM_CHANGER);
    printguid(GUID_DEVCLASS_MODEM);
    printguid(GUID_DEVCLASS_MONITOR);
    printguid(GUID_DEVCLASS_MOUSE);
    printguid(GUID_DEVCLASS_MTD);
    printguid(GUID_DEVCLASS_MULTIFUNCTION);
    printguid(GUID_DEVCLASS_MULTIPORTSERIAL);
    printguid(GUID_DEVCLASS_NET);
    printguid(GUID_DEVCLASS_NETCLIENT);
    printguid(GUID_DEVCLASS_NETSERVICE);
    printguid(GUID_DEVCLASS_NETTRANS);
    printguid(GUID_DEVCLASS_NODRIVER);
    printguid(GUID_DEVCLASS_PCMCIA);
    printguid(GUID_DEVCLASS_PNPPRINTERS);
    printguid(GUID_DEVCLASS_PORTS);
    printguid(GUID_DEVCLASS_PRINTER);
    printguid(GUID_DEVCLASS_PRINTERUPGRADE);
    printguid(GUID_DEVCLASS_PROCESSOR);
    printguid(GUID_DEVCLASS_SBP2);
    printguid(GUID_DEVCLASS_SCSIADAPTER);
    printguid(GUID_DEVCLASS_SECURITYACCELERATOR);
    printguid(GUID_DEVCLASS_SMARTCARDREADER);
    printguid(GUID_DEVCLASS_SOUND);
    printguid(GUID_DEVCLASS_SYSTEM);
    printguid(GUID_DEVCLASS_TAPEDRIVE);
    printguid(GUID_DEVCLASS_UNKNOWN);
    printguid(GUID_DEVCLASS_USB);
    printguid(GUID_DEVCLASS_VOLUME);
    printguid(GUID_DEVCLASS_VOLUMESNAPSHOT);
    printguid(GUID_DEVCLASS_WCEUSBS);

    /* Device class guids for file system filters */
    printguid(GUID_DEVCLASS_FSFILTER_ACTIVITYMONITOR);
    printguid(GUID_DEVCLASS_FSFILTER_ANTIVIRUS);
    printguid(GUID_DEVCLASS_FSFILTER_CFSMETADATASERVER);
    printguid(GUID_DEVCLASS_FSFILTER_COMPRESSION);
    printguid(GUID_DEVCLASS_FSFILTER_CONTENTSCREENER);
    printguid(GUID_DEVCLASS_FSFILTER_CONTINUOUSBACKUP);
    printguid(GUID_DEVCLASS_FSFILTER_COPYPROTECTION);
    printguid(GUID_DEVCLASS_FSFILTER_ENCRYPTION);
    printguid(GUID_DEVCLASS_FSFILTER_HSM);
    printguid(GUID_DEVCLASS_FSFILTER_INFRASTRUCTURE);
    printguid(GUID_DEVCLASS_FSFILTER_OPENFILEBACKUP);
    printguid(GUID_DEVCLASS_FSFILTER_PHYSICALQUOTAMANAGEMENT);
    printguid(GUID_DEVCLASS_FSFILTER_QUOTAMANAGEMENT);
    printguid(GUID_DEVCLASS_FSFILTER_REPLICATION);
    printguid(GUID_DEVCLASS_FSFILTER_SECURITYENHANCER);
    printguid(GUID_DEVCLASS_FSFILTER_SYSTEM);
    printguid(GUID_DEVCLASS_FSFILTER_SYSTEMRECOVERY);
    printguid(GUID_DEVCLASS_FSFILTER_UNDELETE);

#define printstructoffset(s_, f_) do { \
        struct s_ s; \
        if (wanted == NULL || !stricmp(wanted, # s_)) \
            printf("%s(%d).%s: %d bytes at offset %d\n", # s_, sizeof(struct s_), # f_, sizeof(s.f_), offsetof(struct s_, f_)); \
    } while (0)

#if 0
    PRINTNL;
    printstructoffset(_PARTITION_INFORMATION,StartingOffset);
    printstructoffset(_PARTITION_INFORMATION,PartitionLength);
    printstructoffset(_PARTITION_INFORMATION,HiddenSectors);
    printstructoffset(_PARTITION_INFORMATION,PartitionNumber);
    printstructoffset(_PARTITION_INFORMATION,PartitionType);
    printstructoffset(_PARTITION_INFORMATION,BootIndicator);
    printstructoffset(_PARTITION_INFORMATION,RecognizedPartition);
    printstructoffset(_PARTITION_INFORMATION,RewritePartition);

    PRINTNL;
    printstructoffset(_DRIVE_LAYOUT_INFORMATION_EX,PartitionStyle);
    printstructoffset(_DRIVE_LAYOUT_INFORMATION_EX,PartitionCount);
    printstructoffset(_DRIVE_LAYOUT_INFORMATION_EX,Mbr);
    printstructoffset(_DRIVE_LAYOUT_INFORMATION_EX,Gpt);
    printstructoffset(_DRIVE_LAYOUT_INFORMATION_EX,PartitionEntry);

    PRINTNL;
    printstructoffset(_PARTITION_INFORMATION_EX,PartitionStyle);
    printstructoffset(_PARTITION_INFORMATION_EX,StartingOffset);
    printstructoffset(_PARTITION_INFORMATION_EX,PartitionLength);
    printstructoffset(_PARTITION_INFORMATION_EX,PartitionNumber);
    printstructoffset(_PARTITION_INFORMATION_EX,RewritePartition);
    printstructoffset(_PARTITION_INFORMATION_EX,Mbr);
    printstructoffset(_PARTITION_INFORMATION_EX,Gpt);

    PRINTNL;
    printstructoffset(_PARTITION_INFORMATION_GPT, PartitionType);
    printstructoffset(_PARTITION_INFORMATION_GPT, PartitionId);
    printstructoffset(_PARTITION_INFORMATION_GPT, Attributes);
    printstructoffset(_PARTITION_INFORMATION_GPT, Name);

    PRINTNL;
    printstructoffset(_PARTITION_INFORMATION_MBR, PartitionType);
    printstructoffset(_PARTITION_INFORMATION_MBR, BootIndicator);
    printstructoffset(_PARTITION_INFORMATION_MBR, RecognizedPartition);
    printstructoffset(_PARTITION_INFORMATION_MBR, HiddenSectors);

    PRINTNL;
    printstructoffset(_VOLUME_DISK_EXTENTS, NumberOfDiskExtents);
    printstructoffset(_VOLUME_DISK_EXTENTS, Extents);
    printstructoffset(_DISK_EXTENT, DiskNumber);
    printstructoffset(_DISK_EXTENT, StartingOffset);
    printstructoffset(_DISK_EXTENT, ExtentLength);
    return 0;

#endif
}
