# 机票航段整理工具

Windows 离线桌面程序：从 Excel 中识别票务记录，自动拆分往返/联程航段，匹配机场坐标并计算大圆距离，最后输出可由 WPS/Excel 打开的 `.xlsx`。

## 使用方法

1. 从 GitHub Actions 的最新成功构建中下载 `flight-distance-tool-windows`。
2. 解压后双击 `机票航段整理工具.exe`。
3. 选择原始 Excel，填写行程年份，点击“开始整理”。
4. 结果保存在原表同一文件夹，文件名末尾为 `_航段整理结果.xlsx`。

程序不计算碳排放量，也不会上传票务数据。完整机场库已经打包进 EXE，运行时无需联网或安装 Python。

## 支持的记录示例

```text
01.05 AIRTICKET SAMPLE/PERSON VFA-HRE 06JAN
12.23 AIRTICKET TEST/USER VFA-TFU-VFA 12FEB-05MAR
```

无法识别或缺日期的记录会进入“异常复核”工作表，不会静默丢弃。
