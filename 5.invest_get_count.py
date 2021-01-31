import win32com.client
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
print(instCpStockCode.GetCount())


print("--------------------------------")

# 0번째 인덱스에 위치하는 종목의 종목명 확인
print(instCpStockCode.GetData(0, 0))


print("--------------------------------")

# 반복문을 통해 0~9 인덱스에 해당하는 종목명 출력
for i in range(10):
    print(instCpStockCode.GetData(1, i))


print("--------------------------------")

# 네이버 찾기
stockNum = instCpStockCode.GetCount()
for i in range(stockNum):
    if instCpStockCode.Getdata(1, i) == "NAVER":
        print(instCpStockCode.Getdata(0, i))
        print(instCpStockCode.Getdata(1, i))
        print(i)


print("--------------------------------")

# NameToCode, CodeToIndex
naverCode = instCpStockCode.NameToCode("NAVER")
naverIndex = instCpStockCode.CodeToIndex(naverCode)
print(naverCode)
print(naverIndex)