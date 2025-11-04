from spire.xls import *
from spire.xls.common import *
import os

# 创建输出目录
os.makedirs("output/images", exist_ok=True)

# 载入Excel文件
workbook = Workbook()
workbook.LoadFromFile("examples/genexus.xlsx")

# 设置转换器的高画质参数
converterSetting = workbook.ConverterSetting
converterSetting.XDpi = 500  # 水平分辨率 (默认96, 设置300提高画质)
converterSetting.YDpi = 500  # 垂直分辨率 (默认96, 设置300提高画质)
# 如果输出为JPEG格式，可以设置质量（1-100）
# converterSetting.JPEGQuality = 100

# 遍历所有工作表
for i in range(workbook.Worksheets.Count):
    sheet = workbook.Worksheets.get_Item(i)
    
    # 去除白边
    sheet.PageSetup.TopMargin = 0
    sheet.PageSetup.BottomMargin = 0
    sheet.PageSetup.LeftMargin = 0
    sheet.PageSetup.RightMargin = 0
    
    # 将工作表转换为图片
    image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn)
    
    # 保存为PNG文件，文件名包含工作表名称
    image.Save(f"output/images/{sheet.Name}.png")
    print(f"已保存: {sheet.Name}.png")

workbook.Dispose()
print("所有工作表已转换完成！")
