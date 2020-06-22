# Python操作Word文档

## 安装库

docx：https://pypi.org/project/python-docx/

```python
pip install python-docx
```

## 文档操作

### 导入库

```python
from docx import Document
```

### 新建文档

```python
document = Document()
document.save('new.docx')
```

### 打开文档
```python
document = Document('new.docx')
document.save('open.docx')
```

### 写入标题

```python
heading0 = document.add_heading('文档标题', 0)
heading1 = document.add_heading('一级标题', level=1)
heading2 = document.add_heading('二级标题', level=2)
heading3 = document.add_heading('三级标题', level=3)
```

### 写正文
```python
paragraph1 = document.add_paragraph('这是一个段落')
paragraph2 = heading2.insert_paragraph_before('在指定位置之前，插入一个段落')
```

### 设置字体格式
```python
paragraph0 = heading1.insert_paragraph_before('进行字体设置，部分字体')
paragraph0.add_run('加粗').bold = True
paragraph0.add_run('，部分字体')
paragraph0.add_run('倾斜').italic = True
```

### 设置段落格式
```python
document.add_paragraph('引用段落', style='Intense Quote')
document.add_paragraph('插入项目符号', style='List Bullet')
document.add_paragraph('插入编号', style='List Number')
```

### 插入图片
```python
document.add_picture('picture.jpg', width=Inches(5))
```

### 插入分页符
```python
document.add_page_break()
```

### 插入表格
```python
records = (
    (3, '101', 'Spam'),
    (7, '422', 'Eggs'),
    (4, '631', 'Spam, spam, eggs, and spam')
)

table = document.add_table(rows=1, cols=3, style="Light Shading Accent 1")
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Qty'
hdr_cells[1].text = 'Id'
hdr_cells[2].text = 'Desc'
for qty, id, desc in records:
    row_cells = table.add_row().cells
    row_cells[0].text = str(qty)
    row_cells[1].text = id
    row_cells[2].text = desc

document.add_paragraph('更换表格样式')

table2 = document.add_table(rows=1, cols=3, style="Light List Accent 5")
hdr_cells = table2.rows[0].cells
hdr_cells[0].text = 'Qty'
hdr_cells[1].text = 'Id'
hdr_cells[2].text = 'Desc'
for qty, id, desc in records:
    row_cells = table2.add_row().cells
    row_cells[0].text = str(qty)
    row_cells[1].text = id
    row_cells[2].text = desc
```

在这里，你可以手动把word所有的表格样式输出

```python
from docx import Document
from docx.enum.style import WD_STYLE_TYPE

document = Document()
styles = document.styles
 
#生成所有表样式
for s in styles:
    if s.type == WD_STYLE_TYPE.TABLE:
        document.add_paragraph("表格样式 :  "+ s.name)
        table = document.add_table(3,3, style = s)
        heading_cells = table.rows[0].cells
        heading_cells[0].text = '第一列内容'
        heading_cells[1].text = '第二列内容'
        heading_cells[2].text = '第三列内容'
        document.add_paragraph("\n")
 
document.save('Word所有表格样式.docx')
```

具体输出结果，可参考：

### 保存
```
document.save('example.docx')
```

## 效果图

完整代码可参考：

![](http://cdn.zhaojingyi0126.com/IMG/image-20200622134413147.png)

![](http://cdn.zhaojingyi0126.com/IMG/image-20200622135032223.png)