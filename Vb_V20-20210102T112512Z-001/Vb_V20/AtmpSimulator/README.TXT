本程序有两个用途，其一是仿真ATMP，用于测试ATMC，通过自动或手动返回应答信息，测试ATMC的流程
是否正确，尤其是对拒绝码的处理以及对主机控制命令的处理；其二是仿真ATMC，用于测试ATMP，观察
ATMP的返回信息是否与ATMC匹配。


配置文件：
1)PARAM.CFG
记录屏幕上的参数设置
TXKEY=12345678    ATMC与ATMP报文传送中用于DES或SHUFFLE的传输密钥，必须8位或16位
ATMPPORT=1001     本方作为SERVER的TCP/IP端口
ATMCPORT=1001     对方作为SERVER的TCP/IP端口(如ATMC接收主机命令的端口，或者测试对方ATMP的端口)
ATMCIP=127.0.0.1  对方作为SERVER的IP地址
UPDTK=0           如果交易中需要自动更新传输密钥，则需选中
AUTOREPLY=1       如需ATMP自动应答对方的交易请求，则需选中
INCRESP=0         是否取RESPONSE CODE列表中的下一个应答报文


2)ATMP.CFG
仿真ATMP应答的报文格式，根据ATMC发送的报文，组成合适的应答报文

格式说明：(以下报文为交通银行当前报文实例)
应答包格式以“|”分隔，第1段标明ATMC发送来的报文格式(或者是指定RESPONSE CODE的格式，见后)，
其后为应答报文的组成方式，最后以#结束(如#在行首，则整行忽略)

比如：RQK|ABP|YYMMDD|HHMM|YYDDD|SHF:KEY|#
即对ATMC发来的RQK包，返回AAP(03013122550331)(需SHUFFLE的传输密钥)这样的应答包，
其中YYMMDD是年年月月日日格式，HHMM是时时分分的格式，YYDDD是年年日日日的格式

又如：CWD|AQP|ORG:198:20|YYMMDD|HHMM|YYDDD|0000000|SHF:100000|ORG:218:12|ORG:41:141|FIL:0:20|#
即对ATMC发来的CWD包，返回AQP(取自CWD发送包的20位帐号)(日期时间)(7个0)(需SHUFFLE的6位余额)
(取自CWD发送包的12位金额)(取自CWD发送包的104+37位三磁二磁信息)(填20个0)

又如：ATP|ATP|YYMMDD|HHMM|YYDDD|0|SHF:REJ|9901234567890123456789|#
这是指定RESPONSE CODE的报文格式，第1段中的ATP指定该应答包是ATP包，其格式是：
ATP(日期时间)(补个0)(需SHUFFLE的RESPONSE CODE)(填上9901234567890123456789)


3)HOSTCMD.CFG
仿真主机命令发送的报文格式，以及仿真ATMC发送的报文格式

例一：MEC|A1234|#
需发送MECMEC(A1111)，即对1234号ATM进行关机的主机控制命令

例二：ABCDEA1234INQ|FIL:0:26|O|FIL:0:104|40551230114567890=156|FIL:0:16|=502<76336=8>=22|A019?2<8276:=0839;AA|#
需发送的INQ包格式为：(ABCDE包头)(A1234设备号)(INQ)(设备状态等等填26个0)(O表示正在交易)
(三磁信息填104个0)(共37位二磁信息)(16位长的PINBLOCK)(SHUFFLE后的20位卡号)


4)REJCODE.CFG
指定ATMP的应答报文，本方法适用于有多种拒绝码，或者多条卡表文件下载的场合

例一：3003 (等效于3003:ATP)
(同时在ATMP.CFG中有：ATP|ATP|YYMMDD|HHMM|YYDDD|0|SHF:REJ|9901234567890123456789|#)
当RESPONSE CODE列表选定为3003时，应答拒绝码为3003的ATP包，其格式如下：
ATP(日期时间)(补个0)(将列表框数据即3003进行SHUFFLE)(填上9901234567890123456789)

例二：3600:AVP
(同时在ATMP.CFG中有：AUP|AUP|YYMMDD|HHMM|YYDDD|0|SHF:REJ|SHF:150000|01234567890123456789|#)
当RESPONSE CODE列表选定为3600时，应答拒绝码为3600的AUP包，其格式下：
AUP(日期时间)(补个0)(将列表框数据即3600进行SHUFFLE)(SHUFFLE可用余额1500.00)(填20位数字)

例三：
0013060440551200003170410611:ALP (起始的001表示为第一条卡表记录)

9993060460142800003170410611:ALP (起始的999表示为最后一条卡表记录)

(同时在ATMP.CFG中有：ALP|ALP|REJ|#)
当RESPONSE CODE列表选定为0013060440551200003170410611:ALP，且参数AUTOREPLY=1和INCRESP=1时，
该ATMP能够自动回送示例中的两条卡表记录

附：返回包中的关键词表
YYMMDD 6位年年月月日日
YYYYMMDD 8位年年月月日日
YYDDD 5位年年日日日
HHMM 时时分分
REJ 拒绝代码或者其它特定返回(在RESPONSE CODE字段中)
KEY 传输密钥, 8位"12345678"或16位"303132333435363738"格式
FIL:X:n 将字串或字符X重复填充n次, X为可见字符串, 暂不支持chr(i)形式
SHF:X 将后面字串X进行SHUFFLE, X需不超过16字节, 允许采用关键字KEY和REJ
ORG:i:j 从请求包的第i位开始取j字节, 作为应答包的一部分原封不动返回
ORS:i:j 从请求包的第i位开始取j字节, UNSHUFFLE后, 更新传输密钥, 重新SHUFFLE后返回

其它字符将原封不动的加入到返回包中


版本更新信息：
1)2003-03-14
RejectCode栏目保持Enabled
UpdateTK更正一点BUG

