import win32com.client
import pythoncom

def add_rectangle_to_active_slide():
    """
    向当前激活的 PowerPoint 应用程序中的当前选中幻灯片插入一个矩形。
    此函数不依赖文件名，直接操作内存中打开的文档。
    """
    try:
        # 连接到正在运行的 PowerPoint 实例
        # GetActiveObject 获取当前活动的 COM 对象
        powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception as e:
        print("错误: 未检测到正在运行的 PowerPoint 程序。请先打开一个 PPT 文件。")
        print(f"详细信息: {e}")
        return

    try:
        # 获取当前激活的演示文稿
        presentation = powerpoint.ActivePresentation
        
        # 获取当前选中的幻灯片 (View.Slide)
        # 注意：如果处于大纲视图等非普通视图，可能会报错，建议确保在“普通视图”下
        slide = powerpoint.ActiveWindow.View.Slide
        
        # 定义位置和大小 (单位: Points, 1 inch = 72 points)
        left = 72   # 1 inch
        top = 72    # 1 inch
        width = 216 # 3 inches
        height = 144# 2 inches
        
        # 添加矩形形状
        # msoShapeRectangle = 1
        shape = slide.Shapes.AddShape(1, left, top, width, height)
        
        # 设置样式
        shape.Fill.ForeColor.RGB = 255  # 红色 (BGR格式: 0x0000FF -> 255)
        shape.Line.ForeColor.RGB = 16711680  # 蓝色 (BGR格式: 0xFF0000 -> 16711680)
        shape.Line.Weight = 2.5
        
        # print("成功在当前幻灯片中插入矩形！")
        
    except Exception as e:
        print(f"操作失败: {e}")
        print("请确保 PowerPoint 处于打开状态，且当前视图支持编辑幻灯片。")

if __name__ == "__main__":
    add_rectangle_to_active_slide()


def show_shape_range(ppt_app=None):
    """
    显示当前选中幻灯片中的所有形状。
    
    Args:
        ppt_app: PowerPoint 应用程序对象。如果为 None，则尝试获取当前活动的实例。
    """
    try:
        # 如果未提供 ppt_app 对象，则连接到当前运行的 PowerPoint 实例
        if ppt_app is None:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        
        # 获取当前窗口选中的幻灯片范围
        # 注意：AHK中的 Selection.SlideRange 对应 Python 中的 ActiveWindow.Selection.SlideRange
        selection = ppt_app.ActiveWindow.Selection
        slide_range = selection.SlideRange
        
        # 遍历选中的每一张幻灯片
        for slide in slide_range:
            # 检查幻灯片中是否有形状 (Count > 0)
            if slide.Shapes.Count > 0:
                # 获取所有形状的 Range 对象
                # 在 COM 中，Shapes.Range() 不带参数时默认包含所有形状
                shape_range = slide.Shapes.Range()
                
                # 设置可见性为 True (-1 在 VBA/COM 中代表 True)
                shape_range.Visible = True
                
        print("成功显示选中幻灯片中的所有形状！")
        
    except Exception as e:
        print(f"操作失败: {e}")
        print("请确保 PowerPoint 处于打开状态，且已选中至少一张幻灯片。")

if __name__ == "__main__":
    show_shape_range()