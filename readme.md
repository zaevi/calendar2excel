## Calendar2Excel - 生成空Excel日历

![preview](https://user-images.githubusercontent.com/12966814/43789296-29b8bf32-9aa2-11e8-89df-a1dc66b03bcb.PNG)

这是一个方便生成Excel日历的脚本, 提供了大量可自定义的内容, 你可以按照自己需求修改参数并生成不同样式的日历.

- 自定义生成何年何月的日历
- 自定义每日跨列/行数
- 自定义每周开始日
- 自定义各单元格样式, 包括背景色, 字体, 对齐等

### 环境

Python 3.6

xlsxwriter

### 自定义

修改`options`参数控制表格生成:

```python
options = {
    "first-weekday": 0,  # 每周以周几开始(0:周一 ~ 6:周日)
    "day-rows": 5,  # 每日占行数 ≥ 1
    "day-cols": 2  # 每日占列数 ≥ 1
}
```

修改`if(main)`结构控制生成的年月和输出路径:

```python
if __name__ == '__main__':
    today = datetime.date.today()
    filename = "%d-%02d.xlsx" % (today.year, today.month)
    generate(today.year, today.month, filename) # 生成当前月 输出YYYY-MM.xlsx

```

#### 修改样式

修改`styles`参数控制输出的表格样式:

```python
styles = {
    # 头部单元格样式
    "header": {"align":"center", "bg_color":"#A6A6A6"},

    # 针对每日整个区块的样式 允许以list形式轮流变换样式
    "day-block": [{"bg_color":"#C8C8C8"},{"bg_color":"#D9D9D9"}],

    # 每日头部部分的样式, 重复项会覆盖day-block
    "day-header": {"align":"left"},   

    # 每日内容部分的样式, 重复项会覆盖day-block
    "day-text": {},

    # 非本月部分的样式
    "blank-day": {"align":"center", "valign":"vcenter", "bg_color":"#EAEAEA"}, 
}
```
样式参考请参照: [XlsxWriter-Format方法和属性](https://xlsxwriter.readthedocs.io/format.html#format-methods-and-format-properties)

修改示例:

```python
# 本月每周前5天背景色较深
styles["day-block"] = [{"bg_color":"#C8C8C8"}]*5+[{"bg_color":"#D9D9D9"}]*2
styles["blank-day"] = {"align":"center", "valign":"vcenter", "bg_color":"#EAEAEA"}
```

```python
# 头部字体设为12(默认11) 每日内容设为斜体
styles["header"]["font_size"] = 12
styles["day-header"]["font_size"] = 12
styles["day-text"]["italic"] = True
```

