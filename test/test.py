import binascii as b

def hex2str(hex):
    return b.a2b_hex(hex).decode('utf-8')

'''
去掉字符串前后的空格、制表符
'''
def strStrip(s):
    return str(s).strip('\t').strip(' ')


'''
解析HBD请求得到的响应报文
'''
def parseHBD(value):
    value = strStrip(value)
    msgHeader = hex2str(value[0: 8])
    reportMask = value[8: 10]
    length = int(value[10: 12], 16)
    diviceType = value[12: 14]
    protocolVersion = str(value[14: 16])+"."+str(value[16: 18])
    firmwareVersion = str(value[18: 20]) + str(value[20: 22])

    bReportMask = bin(int(reportMask, 16)).replace('0b', '')
    if bReportMask[3] == 0:
        uniqueID = ''
        for i in range(8):
            index = 22+i*2
            uniqueID += str(int(value[index: index+2], 16))
    else:
        uniqueID = hex2str(value[22: 38])
    sendTime = str(int(value[38: 42], 16)) + str(int(value[42: 43], 16)) + str(int(value[43: 44], 16)) + str(int(value[44: 46], 16)) \
               + str(int(value[46: 47], 16)) + str(int(value[47: 48], 16)) + str(int(value[48: 50], 16)) + str(int(value[50: 52], 16))
    countNumber = value[52: 56]
    checksum = value[56: 60]
    tailCharacters = value[60: 64]
    result = "Message Header: "+msgHeader + "\nReport Mask: " + reportMask + "\nLength: " +\
    str(length) +"\nDevice Type: " + diviceType + "\nProtocol Version: " + protocolVersion + "\nFirmware Version: " +\
    firmwareVersion + "\nUnique ID: " + uniqueID + "\nSend Time: " + sendTime + "\nCount Number: " +\
    countNumber + "\nChecksum: " + checksum + "\nTail Characters: " + tailCharacters
    print(result)
    return result

'''
解析DAT请求得到的响应报文
'''
def parseDAT(value):
    value = strStrip(value)
    print(value)

'''
解析EVT请求得到的响应报文
'''
def parseEVT(value):
    value = strStrip(value)
    msgHeader = hex2str(value[0: 8])
    reportMask = value[8: 10]
    length = int(value[10: 12], 16)
    diviceType = value[12: 14]
    protocolVersion = str(value[14: 16]) + "." + str(value[16: 18])
    firmwareVersion = str(value[18: 20]) + str(value[20: 22])

'''
解析请求以及响应报文
'''
def parse(key, value):
    key = strStrip(key)
    key = key[1:4]
    resultFile = open('data/result.log','w')
    if 'HBD' == key:
        resultFile.write(parseHBD(value))
    elif 'EVT' == key:
        parseEVT(value)

    resultFile.flush()
    resultFile.close()


f = open("data/GV300N.log", 'r')
aesc = 0  # 区分报文日志的顺序，1代表正序，0代表倒序
hashMap = {}
value = ""

while True:
    line = f.readline()
    if not line:
        break
    # print(line)
    if aesc == 1:  # 顺序
        line = f.readline()
    else:  # 倒序
        if line.__contains__('HEX'):
            value = line[line.find('HEX:') + len("HEX:"):]
            # print(value)
        elif line.__contains__('ASC'):
            key = line[line.find('ASC:') + len("ASC:"):]
            # print(key)
            hashMap[key] = value


# print(count)
for key, value in hashMap.items():
    # print('key is %s,value is %s' % (key, value))
    parse(key, value)
