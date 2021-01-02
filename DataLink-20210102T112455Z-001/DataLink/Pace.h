
//Define the 25 field seperators in IBM Pace format 
///////////////////////////////////////////////////
//FS01 ---- Request Message Head ID
#define	MESGHEADID               { 0x27,0x20,0x20 }
//FS02 ---- Message Type FS
#define MESGIDFS                 { 0x11,0x20,0x41 }
//FS03 ---- Message Sequence NO. FS
#define MESGSEQFS                { 0x11,0x20,0x49 }
//FS04 ---- Logical unit number FS
#define MESGLUNOFS               { 0x11,0x20,0x26 }
//FS05 ---- Customer Card No. FS
#define TREQCARDNOFS             { 0x11,0x20,0x4F }
//FS06 ---- Transaction code FS
#define TREQTRANSCODEFS          { 0x11,0x20,0x5F }
//FS07 ---- Transaction currency FS
#define TREQTRANSCURRFS          { 0x11,0x20,0x36 }
//FS08 ---- Transaction amount FS
#define TREQTRANAMOUNTFS         { 0x11,0x20,0x37 }
//FS09 ---- Transaction serial number FS
#define TREQSEQNUMFS             { 0x11,0x41,0x42 }
//FS10 ---- Transaction time date FS
#define TREQTIMEDATEFS           { 0x11,0x41,0x4A }
//FS11 ---- Transaction status FS
#define TREQTRANSTATFS           { 0x11,0x41,0x49 }
//FS12 ---- Transaction PIN FS
#define TREQPINFS                { 0x11,0x41,0x3B }
//FS13 ---- Transaction Track2 FS
#define TREQTRACK2FS             { 0x11,0x42,0x2F }
//FS14 ---- Transaction Track3 FS
#define TREQTRACK3FS             { 0x11,0x43,0x31 }
//FS15 ---- Transaction Customer Input FS
#define TREQINPUTBUFFS           { 0x11,0x2E,0x5B }
//FS16 ---- Transaction Misc data FS
#define TREQMISCDATAFS           { 0x11,0x2E,0x52 }
//FS17 ---- Transaction IC Card data FS
#define TREQICDATAFS             { 0x11,0x20,0x55 }
//FS18 ---- Complete trans serial FS
#define COMPSEQNUMFS             { 0x11,0x41,0x42 }
//FS19 ---- Complete trans code FS
#define COMPTRANSCODEFS          { 0x11,0x41,0x49 }
//FS20 ---- Reject Code FS
#define REJTCODEFS               { 0x11,0x20,0x4F }
//FS21 ---- Stat2 FS
#define STA2FS                   { 0x11,0x20,0x4F }
//FS22 ---- Response Host command return code FS
#define RESPHOSTCMDRETCODEFS     { 0x11,0x20,0x4F }
//FS23 ---- Response Host command ID FS
#define RESPHOSTCMDCODEFS        { 0x11,0x20,0x29 }
//FS24 ---- Response Host command return data FS
#define RESPHOSTCMDRETDATAFS     { 0x11,0x20,0x55 }
//FS25 ---- Transaction MAC FS
#define MESGMACFS                { 0x11,0x2E,0x53 }