import xlwt

ReadPath = r"./input.txt"
SavePath = r"./AWROutput.xls"
anti_Accuracy_Value = 0         #相隔多少个数据删除一个


def anti_Accuracy(Input):
    if anti_Accuracy_Value < 1:
        return Input
    l = len(Input)
    Output = []
    Output.append(Input[0])
    i = 1
    while i < l:
        Output.append(Input[i])
        i = i+anti_Accuracy_Value+1
    return Output

def ReadAWROutput():
    fileOpen = open(ReadPath)
    DataBuffer = fileOpen.readlines()
    fileOpen.close()
    for i in range(0, len(DataBuffer)):
        DataBuffer[i] = DataBuffer[i].split('\t')

        for j in range(0, len(DataBuffer[i])):
            try:
                DataBuffer[i][j] = float(DataBuffer[i][j])
            except Exception as results:
                print("转换数字错误" + str(results))

    return DataBuffer


def SaveData2Excel(Datas):
    workbook = xlwt.Workbook(encoding="utf-8")
    worksheet = workbook.add_sheet('AWR Output', cell_overwrite_ok=True)

    for j in range(0, len(Datas)):
        data = Datas[j]
        for k in range(0, len(data)):
            worksheet.write(j + 1, k, data[k])
    try:
        workbook.save(SavePath)
        print("已将请求数据写入Excel表格")
    except Exception as results:
        print("Excel写入错误: " + str(results))
        return 1


def main():
    SaveData2Excel(anti_Accuracy(ReadAWROutput()))


if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
