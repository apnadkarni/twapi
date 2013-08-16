#include "twapi.h"
#include "lzmadec.h"

/* Allocators needed by the LZMA library */
static void *TwapiLzmaAlloc(void *unused, size_t size) 
{
    return TwapiAlloc(size); 
}

static void TwapiLzmaFree(void *unused, void *address) 
{
    TwapiFree(address);
}

/* Struct defined by LZMA library */
static ISzAlloc gLzmaAlloc = { TwapiLzmaAlloc, TwapiLzmaFree };

void TwapiLzmaFreeBuffer(unsigned char *buf)
{
    if (buf)
        TwapiFree(buf);
}

unsigned char *TwapiLzmaUncompressBuffer(Tcl_Interp *interp,
                                         unsigned char *indata,
                                         DWORD insz, DWORD *outszP)
{
    unsigned char *outdata = NULL;
    /* header: 5 bytes of LZMA properties and 8 bytes of uncompressed size */
    int i;
    SRes res = 0;
    UInt64 outsz;
    SizeT inlen, outlen;
    ELzmaStatus status;

    if (insz < (LZMA_PROPS_SIZE+8)) {
        ObjSetStaticResult(interp, "Input LZMA data header too small.");
        return NULL;
    }

    /* Figure out the uncompressed size from the LZMA header */
    outsz = 0;
    for (i = 0; i < 8; i++)
        outsz += (UInt64)indata[LZMA_PROPS_SIZE + i] << (i * 8);

    if (outsz == 0xffffffffffffffff) {
        ObjSetStaticResult(interp, "No length field in LZMA data. Propably compressed with eos marker. This is not supported.");
        return NULL;
    }

    outdata = TwapiAlloc((size_t) outsz);
    
    outlen = (SizeT) outsz;
    inlen = insz - LZMA_PROPS_SIZE - 8;
    res = LzmaDecode(outdata, &outlen,
                     indata+LZMA_PROPS_SIZE+8, &inlen, /* Actual data */
                     indata, LZMA_PROPS_SIZE, /* LZMA properties */
                     LZMA_FINISH_END, &status,
                     &gLzmaAlloc);

    if (res != SZ_OK ||
        outlen != outsz ||
        inlen != (insz - LZMA_PROPS_SIZE - 8) ||
        (status != LZMA_STATUS_FINISHED_WITH_MARK &&
         status != LZMA_STATUS_MAYBE_FINISHED_WITHOUT_MARK)) {
        goto error_return;

    }
    
    *outszP = (DWORD) outlen;
    return outdata;

error_return:
    ObjSetStaticResult(interp, "LzmaDecode failed.");
    if (outdata)
        TwapiFree(outdata);
    return NULL;
}

