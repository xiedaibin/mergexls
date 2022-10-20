import os
import xlwings as xw

# 定义个文件操作类
class FileTable(object):
    def __init__(self, name, header,data,rows,cols):
        self.name = name
        self.header = header
        self.data = data
        self.rows = rows
        self.cols = cols
        

# 遍历文件夹获取文件
def file_name_walk(file_dir):
    for root, dirs, files in os.walk(file_dir):
        #print("root", root)  # 当前目录路径
        #print("dirs", dirs)  # 当前路径下所有子目录
        #print("files", files)  # 当前路径下所有非目录子文件
        return files

#获取表数据
def get_table_datas(app,files,data_path):
    
    tableList = []
    for file in files:
        fileExt = os.path.splitext(file)[-1]
        if fileExt == ".xls" or fileExt == ".xlsx":
            xf = os.path.join(data_path,file)
            wb = app.books.open(xf)
            sheet = wb.sheets[0]
            info = sheet.used_range
            nrows = info.last_cell.row
            ncols = info.last_cell.column
            print("可用行:",nrows," 可用列",ncols)
            #获取第一行 表头数据数据
            a1_an_value = sheet.range((1,1),(1,ncols)).value
            #print("表头数据:",a1_an_value)
            a1_nn_value = sheet.range((1,1),(nrows,ncols)).value
            #print("表格数据:",a1_nn_value)
            fileTable = FileTable(file,a1_an_value,a1_nn_value,nrows,ncols)
            tableList.append(fileTable)
            wb.close()
    return tableList

#获取合并新表头数据
def get_header_datas(tableDatas):
    headerList = []
    # 合并表头
    for td in tableDatas:
        headerList.extend(td.header)
    #去重
    headerList = list(set(headerList))
    #排序
    headerList.sort()
    return headerList

# main 主函数开始
cpath,filename = os.path.split(os.path.abspath(__file__))

print("执行文件目录:"+cpath+" 执行文件:"+filename)
# 切换为测试数据 将data 改为 test 则会合并测试目录的数据
data_path = os.path.join(cpath,"data")
#data_path = os.path.join(cpath,"test")
print("数据路径:"+data_path)
   
files = file_name_walk(data_path)
print("需要处理得文件:", files)
app = xw.App(visible=False,add_book=False)
app.screen_updating = False # 关闭刷新，excel不显示打开的表格内容，可以少许提高速度，但如果长时间停留在这种状态会造成excel失去响应的假象
app.display_alerts = False
#app.visible = False # 可以不显示excel界面，但会闪现一下 初始化设置将不会闪现
tableDatas = get_table_datas(app=app,files=files,data_path=data_path)
headerList = get_header_datas(tableDatas)
print("表头合并后数据量:", len(headerList))

#新建文档并写入表头
sabeWb =app.books.add() 
saveSheet = sabeWb.sheets[0]
# 填充表头
saveSheet.range('A1').value = headerList
# 填充数据
# 遍历数据表
# 从第二行开始插入数据
newRowIndex = 2
totalData = 1
# 用二维宿主保存数据再统一保存到excel
for td in tableDatas:
    totalData = totalData + (td.rows-1)
    matrix = [['' for col in range(len(headerList))] for row in range(td.rows-1)]
    #print("初始二维数组:",matrix)
    #遍历数据 从第二行开始遍历，第一行为表头
    beginRowIndex = newRowIndex
    for rowIndex,row in enumerate(td.data[1:]):
        for index in range(len(td.header)):
            #print(td.header[index],'列 第',rowIndex+2,'行 cell数据:',row[index])
            # 列下标 数组下标从0开始
            newHeaderIndex = headerList.index(td.header[index])
            #print("新列 Index",newHeaderIndex)
            matrix[rowIndex][newHeaderIndex] = row[index]
        #print(td.name,'行数据:',row)
        newRowIndex = newRowIndex + 1
    beginRowIndexStr = 'A' +str(beginRowIndex)
    saveSheet.range(beginRowIndexStr).value = matrix
    #print("表赋值后二维数组:",matrix)
    print("表",td.name,"插入新表成功")
savePath = os.path.join(data_path,'save')
# 创建保存文件路径
if not os.path.exists(savePath):
    os.makedirs(savePath)
saveFile = os.path.join(savePath,'merge.xls')
sabeWb.save(saveFile)
sabeWb.close()
app.kill()
print("合并完成,总数据量为:",totalData)

