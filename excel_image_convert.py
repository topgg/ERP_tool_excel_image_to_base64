import time
import base64
from io import BytesIO
from PIL import ImageGrab
import win32com.client as win32
import os

# 输入文件路径
input_file = r'E:\供应商管理\照明厂商\价格维护表 - 副本.xlsx'

# Excel应用和工作簿设置
excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open(input_file)

for sheet in workbook.Worksheets:
    for shape in list(sheet.Shapes):  # 使用list来避免在迭代期间修改集合的错误
        if shape.Name.startswith('Picture'):
            shape.Copy()
            time.sleep(0.3)
            image = ImageGrab.grabclipboard()
            
            # 将图片转换为Base64
            buffered = BytesIO()
            if image.mode != 'RGB':
                image = image.convert('RGB')
            image.save(buffered, format="JPEG",quality=70)
            img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
            
            # 将Base64编码插入到Excel单元格中
            cell = shape.TopLeftCell
            cell.Value = img_base64
            
            # 删除原始图片
            shape.Delete()

# 保存为新文件
new_file_name = "new" + os.path.basename(input_file)
new_file_path = os.path.join(os.path.dirname(input_file), new_file_name)
workbook.Close()
excel.Quit()
