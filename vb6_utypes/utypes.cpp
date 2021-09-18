#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <math.h>
#include <cstdlib>

#define CHECK_BIT(var,pos) (((var)>>(pos)) & 1)

enum op{
	op_add = 0,
	op_sub = 1,
	op_div = 2,
	op_mul = 3,
	op_mod = 4,
	op_xor = 5,
	op_and = 6,
	op_or  = 7,
	op_rsh = 8,
	op_lsh = 9,
	op_gt  = 10,
	op_lt  = 11,
	op_gteq = 12,
	op_lteq = 13,
	op_rol32 = 14,
	op_setBit = 15,
	op_clearBit = 16,
	op_toggleBit = 17,
	op_testBit = 18
};

enum modes{
	mUnsigned = 0,
	mSigned = 1,
	mHex = 2
};

struct x64{
	unsigned int lo;
	unsigned int hi;
};

double __stdcall entropy(unsigned char* buf, unsigned int length)
{
	double temp;
	double bytes[256] = {0.0};
	double e = 1.4426950408889634073599246810023;

	for(int i = 0; i < length; i++){
		unsigned char c = buf[i];
		bytes[c] += 1;
	}

    for(int j = 0; j < 256; j++){
        temp = bytes[j] / (double)length;
        if(temp) e += ( -log(temp) / log((double)2) ) * bytes[j];
    }

    e = e /(double)length;
    return e;
}

void __stdcall vc_srand(unsigned int v1){srand(v1);}
unsigned int __stdcall vc_rand(){return rand();}

unsigned int __stdcall ULong(unsigned int v1, unsigned int v2, int operation){

	switch(operation){
		case op_add: return v1 + v2;
		case op_sub: return v1 - v2;
		case op_div: return v1 / v2;
		case op_mul: return v1 * v2;
		case op_mod: return v1 % v2;
		case op_xor: return v1 ^ v2;
		case op_and: return v1 & v2;
		case op_or:  return v1 | v2;
		case op_rsh: return v1 >> v2;
		case op_lsh: return v1 << v2;
		case op_gt: return (v1 > v2 ? 1 : 0);
		case op_lt: return (v1 < v2 ? 1 : 0);
		case op_gteq: return (v1 >= v2 ? 1 : 0);
		case op_lteq: return (v1 <= v2 ? 1 : 0);
		case op_rol32: return _rotl(v1,v2);
		case op_setBit: v1 |= 1 << v2; return v1;
		case op_clearBit: v1 &= ~(1 << v2); return v1; 
		case op_toggleBit: v1 ^= 1 << v2; return v1;
		case op_testBit: return CHECK_BIT(v1,v2);
	}

	return -1;

}

unsigned short __stdcall UInt(unsigned short v1, unsigned short v2, int operation){

	switch(operation){
		case op_add: return v1 + v2;
		case op_sub: return v1 - v2;
		case op_div: return v1 / v2;
		case op_mul: return v1 * v2;
		case op_mod: return v1 % v2;
		case op_xor: return v1 ^ v2;
		case op_and: return v1 & v2;
		case op_or:  return v1 | v2;
		case op_rsh: return v1 >> v2;
		case op_lsh: return v1 << v2;
		case op_gt: return (v1 > v2 ? 1 : 0);
		case op_lt: return (v1 < v2 ? 1 : 0);
		case op_gteq: return (v1 >= v2 ? 1 : 0);
		case op_lteq: return (v1 <= v2 ? 1 : 0);
		case op_rol32: return _rotl(v1,v2);
		case op_setBit: v1 |= 1 << v2; return v1;
		case op_clearBit: v1 &= ~(1 << v2); return v1; 
		case op_toggleBit: v1 ^= 1 << v2; return v1;
		case op_testBit: return CHECK_BIT(v1,v2);
	}

	return -1;

}  

//this is just because i need the C style overflow wrap for a project..
unsigned char __stdcall UByte(unsigned char v1, unsigned char v2, int operation){

	switch(operation){
		case op_add: return v1 + v2;
		case op_sub: return v1 - v2;
		case op_div: return v1 / v2;
		case op_mul: return v1 * v2;
		case op_mod: return v1 % v2;
		case op_xor: return v1 ^ v2;
		case op_and: return v1 & v2;
		case op_or:  return v1 | v2;
		case op_rsh: return v1 >> v2;
		case op_lsh: return v1 << v2;
		case op_gt: return (v1 > v2 ? 1 : 0);
		case op_lt: return (v1 < v2 ? 1 : 0);
		case op_gteq: return (v1 >= v2 ? 1 : 0);
		case op_lteq: return (v1 <= v2 ? 1 : 0);
		case op_rol32: return _rotl(v1,v2);
		case op_setBit: v1 |= 1 << v2; return v1;
		case op_clearBit: v1 &= ~(1 << v2); return v1; 
		case op_toggleBit: v1 ^= 1 << v2; return v1;
		case op_testBit: return CHECK_BIT(v1,v2);
	}

	return -0;

}  

unsigned __int64 __stdcall U64(unsigned __int64 v1, unsigned __int64 v2, int operation){

	switch(operation){
		case op_add: return v1 + v2;
		case op_sub: return v1 - v2;
		case op_div: return v1 / v2;
		case op_mul: return v1 * v2;
		case op_mod: return v1 % v2;
		case op_xor: return v1 ^ v2;
		case op_and: return v1 & v2;
		case op_or:  return v1 | v2;
		case op_rsh: return v1 >> v2;
		case op_lsh: return v1 << v2;
		case op_gt: return (v1 > v2 ? 1 : 0);
		case op_lt: return (v1 < v2 ? 1 : 0);
		case op_gteq: return (v1 >= v2 ? 1 : 0);
		case op_lteq: return (v1 <= v2 ? 1 : 0);
		case op_rol32: return _rotl(v1,v2);
		case op_setBit: v1 |= 1 << v2; return v1;
		case op_clearBit: v1 &= ~(1 << v2); return v1; 
		case op_toggleBit: v1 ^= 1 << v2; return v1;
		case op_testBit: return CHECK_BIT(v1,v2);
	}

	return -1;

}

__int64 __stdcall S64(__int64 v1, __int64 v2, int operation){

	switch(operation){
		case op_add: return v1 + v2;
		case op_sub: return v1 - v2;
		case op_div: return v1 / v2;
		case op_mul: return v1 * v2;
		case op_mod: return v1 % v2;
		case op_xor: return v1 ^ v2;
		case op_and: return v1 & v2;
		case op_or:  return v1 | v2;
		case op_rsh: return v1 >> v2;
		case op_lsh: return v1 << v2;
		case op_gt: return (v1 > v2 ? 1 : 0);
		case op_lt: return (v1 < v2 ? 1 : 0);
		case op_gteq: return (v1 >= v2 ? 1 : 0);
		case op_lteq: return (v1 <= v2 ? 1 : 0);
		case op_rol32: return _rotl(v1,v2);
		case op_setBit: v1 |= 1 << v2; return v1;
		case op_clearBit: v1 &= ~(1 << v2); return v1; 
		case op_toggleBit: v1 ^= 1 << v2; return v1;
		case op_testBit: return CHECK_BIT(v1,v2);
	}

	return -1;

}


int __stdcall U642Str(unsigned __int64 v1, LPSTR pszString, LONG cSize, int mode){

	char buf[64]={0};
    signed __int64 v2 = v1;
    int i;

	switch(mode){
		case mUnsigned: sprintf(buf, "%I64u", v1); break;
		case mSigned:   sprintf(buf, "%I64d", v1); break;
		case mHex:      sprintf(buf, "%I64x", v1); break;
	}
	
	//printf("%08X%08X", static_cast<UINT32>((u64>>32)&0xFFFFFFFF), static_cast<UINT32>(u64)&0xFFFFFFFF));

	i = strlen(buf);
	if (cSize > i ) strcpy(pszString, buf);
	return i;

}

int __stdcall U2Str(unsigned int v1, LPSTR pszString, LONG cSize){

	char buf[64]={0};
    int i;
    sprintf(buf, "%u", v1);
	i = strlen(buf);
	if (cSize > i ) strcpy(pszString, buf);
	return i;

}

int __stdcall UShortStr(unsigned short v1, LPSTR pszString, LONG cSize){

	char buf[64]={0};
    int i;
    sprintf(buf, "%hu", v1);
	i = strlen(buf);
	if (cSize > i ) strcpy(pszString, buf);
	return i;

}

unsigned int  __stdcall Str2Long(LPSTR pszString, int base){
	//supports - but doesnt seem to work with 0x prefix? so had to include base arg..
	unsigned int ret = strtoul(pszString, NULL, base);
	return ret;
}

unsigned __int64  __stdcall Str264(LPSTR pszString, int base){
	//supports - but doesnt seem to work with 0x prefix? so had to include base arg..
	unsigned __int64 ret = _strtoui64(pszString, NULL, base);
	return ret;
}

unsigned __int64 __stdcall toU64(unsigned int v1, unsigned int v2){

	unsigned __int64 ret = 0;
	x64 *x = (struct x64*)&ret;
	x->hi = v1;
	x->lo = v2;
	return ret;

}

void __stdcall fromU64(unsigned __int64 v0, unsigned int *v1, unsigned int *v2){

	x64 *x = (struct x64*)&v0;
	*v1 = x->hi;
	*v2 = x->lo;

}


/* 

was hoping doubles allowed you to do native + - in vb6 but seems wonky..
Private Declare Sub toU64d Lib "utypes.dll" (ByVal v1 As Long, ByVal v2 As Long, ByRef outVar As Double)

void __stdcall toU64d(unsigned int v1, unsigned int v2, unsigned __int64 *outVar){

	unsigned __int64 ret = 0;
	x64 *x = (struct x64*)&ret;
	x->hi = v1;
	x->lo = v2;
	*outVar = ret;

}*/