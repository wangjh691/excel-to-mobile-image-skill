#!/bin/bash
echo "=================================================="
echo "      正在生成技术申请手机长图，请稍候...         "
echo "=================================================="
python3 /Users/wangjunhua/Desktop/00.工作资料/agi/SKILL/excel-to-mobile-image-skill/scripts/excel_to_image.py
echo "=================================================="
echo " 转换完成！已在当前目录生成 手机长图。   "
echo "=================================================="

# 删除过程文档
rm -f "/Users/wangjunhua/Desktop/00.工作资料/agi/工作区/工作安排/(技术管理)技术申请_整理.xlsx"

sleep 3

# 异步延迟 1 秒后自动关闭当前终端窗口，避开有任务运行的警告
(sleep 1; osascript -e 'tell application "Terminal" to close front window') &
exit
