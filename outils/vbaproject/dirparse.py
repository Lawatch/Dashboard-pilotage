import struct
def records(d):
    i=0; out=[]
    while i < len(d):
        rid,size=struct.unpack_from('<HI',d,i)
        if rid==0x0009:  # PROJECTVERSION: Reserved(4) + Major(4) + Minor(2)
            out.append((rid,i,d[i:i+12])); i+=12; continue
        out.append((rid,i,d[i:i+6+size])); i+=6+size
    return out
