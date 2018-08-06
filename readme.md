### Calendar2Excel - 生成空Excel日历

![preview](https://user-images.githubusercontent.com/12966814/43713228-6d809016-99ab-11e8-9408-fc262251d24e.png)

####环境

Python 3.6

xlsxwriter

####自定义

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

