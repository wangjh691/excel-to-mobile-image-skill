#!/usr/bin/env python3
import os
import sys
import argparse
import datetime
import subprocess
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image

# 设置中文字体首选和 CSS 样式
CSS_VARIABLES = """
:root {
  --primary: #0f172a;       /* slate-900 */
  --primary-light: #1e293b; /* slate-800 */
  --accent: #2563eb;        /* blue-600 */
  --accent-light: #eff6ff;  /* blue-50 */
  --bg-main: #f3f5f8;       /* slate-100 偏莫兰迪浅灰蓝，打造通透底色 */
  --bg-card: #ffffff;
  --text-main: #334155;     /* slate-700 柔和文本，降低阅读对比度疲劳 */
  --text-dark: #0f172a;     /* slate-900 */
  --text-muted: #64748b;    /* slate-500 */
  --border-light: #e2e8f0;  /* slate-200 */
  
  /* 统一圆角比例：大圆角 16px，小圆角 6px */
  --radius-lg: 16px;
  --radius-sm: 6px;
  
  /* 莫兰迪低饱和度雅致颜色 */
  --badge-yes-bg: #f0fdf4;   /* 极浅绿 */
  --badge-yes-text: #166534;  /* 深绿 */
  --badge-no-bg: #f1f5f9;    /* 极浅灰 */
  --badge-no-text: #475569;   /* 深灰 */
  
  /* 兵种标签统一为冷色调拼图 */
  --badge-af-bg: #f0f9ff;    /* 安服 - 浅蓝 */
  --badge-af-text: #0369a1;  /* 安服 - 深蓝 */
  --badge-xf-bg: #f0fdfa;    /* 行销 - 极浅青 */
  --badge-xf-text: #0d9488;  /* 行销 - 深青 */
  --badge-jf-bg: #eef2ff;    /* 技服 - 极浅靛 */
  --badge-jf-text: #4f46e5;  /* 技服 - 深靛 */
  --badge-other-bg: #f8fafc; /* 其他 */
  --badge-other-text: #64748b;
  
  /* 物理高光微阴影，混合背景靛蓝色调 (Tinted shadows) */
  --shadow-card: 0 4px 20px rgba(15, 23, 42, 0.02), 0 2px 4px rgba(15, 23, 42, 0.01), inset 0 1px 0 rgba(255, 255, 255, 0.95);
  --shadow-stats: 0 6px 16px rgba(15, 23, 42, 0.04), 0 1px 3px rgba(15, 23, 42, 0.02);
}
"""

def parse_args():
    parser = argparse.ArgumentParser(description="将工作计划 Excel 转换为精美的手机查看图片")
    parser.add_argument("-i", "--input", required=True, help="输入的 Excel 文件路径")
    parser.add_argument("-o", "--output-dir", help="输出文件夹路径（默认为输入文件同级目录）")
    return parser.parse_args()

def find_chrome_binary():
    """
    智能寻找系统中的 Chrome 或 Chromium 可执行文件路径，提高跨平台运行稳定性
    """
    # 1. 尝试从环境变量获取
    env_path = os.environ.get("CHROME_PATH")
    if env_path and os.path.exists(env_path):
        return env_path
        
    # 2. 针对 macOS 的默认路径
    mac_paths = [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        "/Applications/Chromium.app/Contents/MacOS/Chromium"
    ]
    for p in mac_paths:
        if os.path.exists(p):
            return p
            
    # 3. 针对 Linux/Unix 系统的 path 查找
    import shutil
    for cmd in ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "microsoft-edge"]:
        path = shutil.which(cmd)
        if path:
            return path
            
    return None

def clean_and_prepare_data(input_path):
    """
    读取并清洗 Excel 数据，按销售排序，新增序号，保留指定字段
    """
    if not os.path.exists(input_path):
        print(f"错误：输入文件不存在 {input_path}")
        sys.exit(1)
        
    df = pd.read_excel(input_path)
    
    # 1. 排序（按销售）
    # 渐进式拼音排序：优先尝试使用 pypinyin 进行完美汉字拼音排序；未安装时 fallback 到 GB18030 字符字节排序
    if '销售' in df.columns:
        try:
            from pypinyin import pinyin, Style
            def get_pinyin_key(name):
                p_list = pinyin(str(name), style=Style.NORMAL)
                return "".join([w[0] for w in p_list]).lower()
            df = df.sort_values(by='销售', key=lambda col: col.map(get_pinyin_key))
        except ImportError:
            # fallback
            df = df.sort_values(by='销售', key=lambda col: col.map(lambda x: str(x).encode('gb18030', errors='ignore')))
        df = df.reset_index(drop=True)
    else:
        df['销售'] = ''
        
    # 3. 填充 NaN 为空字符串，预计工时如果为 NaN 填 0 或空，保持数据干净
    df = df.fillna('')
    
    # 4. 新增序号字段并放置在第一列
    df.insert(0, '序号', range(1, len(df) + 1))
    
    # 5. 保留指定字段并按顺序排列
    target_columns = ['序号', '所在月周', '销售', '兵种', '客户名称', '工作类型', '是否支持', '支持人员', '计划时间', '预计工时（h）', '备注']
    
    # 容错：如果缺少某些字段，创建空列
    for col in target_columns:
        if col not in df.columns:
            df[col] = ''
            
    df = df[target_columns]
    return df

def beautify_excel(file_path):
    """
    使用 openpyxl 对导出的 Excel 进行精美样式美化
    """
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    font_family = "Microsoft YaHei"
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid") # slate-800
    
    data_font = Font(name=font_family, size=10)
    data_fill_even = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # slate-50
    data_fill_odd = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    thin_border_side = Side(border_style="thin", color="CBD5E1") # slate-300
    border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    
    # 格式化表头
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = border
        
    # 格式化数据行
    for row_idx in range(2, ws.max_row + 1):
        fill = data_fill_even if row_idx % 2 == 0 else data_fill_odd
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.fill = fill
            cell.border = border
            
            # 对齐逻辑
            col_name = ws.cell(row=1, column=col_idx).value
            if col_name in ['序号', '所在月周', '销售', '是否支持', '计划时间']:
                cell.alignment = align_center
            elif col_name in ['预计工时（h）']:
                cell.alignment = align_right
                # 尽量保证数字类型
                try:
                    if cell.value != '':
                        cell.value = float(cell.value)
                except ValueError:
                    pass
            else:
                cell.alignment = align_left
                
    # 自动调整列宽
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            cell_len = sum(2 if '\u4e00' <= c <= '\u9fff' else 1 for c in val)
            if cell_len > max_len:
                max_len = cell_len
        ws.column_dimensions[col_letter].width = max(max_len + 4, 10)
        
    # 行高
    ws.row_dimensions[1].height = 28
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 22
        
    wb.save(file_path)

def crop_image_by_marker(image_path, marker_color=(15, 23, 42), tolerance=12):
    """
    从下往上扫描图片，寻找特定颜色的 marker 横线，进行高度裁剪以防多余留白 and 内容截断
    """
    import os
    if not os.path.exists(image_path):
        return False
    try:
        img = Image.open(image_path)
        width, height = img.size
        
        # 从下往上每隔 2 行扫描一次，大幅提升效率
        # 限制最大扫描深度（向上最深扫描至整体的 20% 高度），防止在超长空白时发生误判并减少计算量
        max_scan_limit = max(200, height // 5)
        marker_y = None
        for y in range(height - 10, max_scan_limit, -2):
            match_count = 0
            # 在图像中间 30% 范围内采样，确保效率并避开边缘渐变
            sample_points = range(int(width * 0.35), int(width * 0.65), 6)
            for x in sample_points:
                pixel = img.getpixel((x, y))
                # 兼容 RGB 和 RGBA 格式
                r, g, b = pixel[0], pixel[1], pixel[2]
                if (abs(r - marker_color[0]) <= tolerance and 
                    abs(g - marker_color[1]) <= tolerance and 
                    abs(b - marker_color[2]) <= tolerance):
                    match_count += 1
            
            # 如果这一行有 80% 的采样像素点符合颜色，说明找到了底部分割线
            if match_count > (len(sample_points) * 0.8):
                marker_y = y
                break
                
        if marker_y is not None:
            # 裁剪到 marker 线下方 10 像素处
            cropped_img = img.crop((0, 0, width, min(height, marker_y + 10)))
            cropped_img.save(image_path)
            print(f"智能图像处理成功：检测到结束标记 (y={marker_y})，长图高度自适应裁剪为 {marker_y + 10}px")
            return True
        else:
            print("提示：未能在图像中定位底部分割线，保持原图大小。")
            return False
    except Exception as e:
        print(f"图像裁剪处理出错：{e}")
        return False

def get_badge_class(val):
    if val == "是":
        return "badge-yes"
    elif val == "否":
        return "badge-no"
    elif "安服" in str(val):
        return "badge-af"
    elif "行销" in str(val):
        return "badge-xf"
    elif "技服" in str(val):
        return "badge-jf"
    else:
        return "badge-other"

def generate_table_html(df, stats):
    """
    生成精美大表格 HTML 字符串
    """
    rows_html = ""
    for _, row in df.iterrows():
        support_cls = get_badge_class(row['是否支持'])
        type_cls = get_badge_class(row['兵种'])
        
        rows_html += f"""
        <tr>
            <td class="text-center">{row['序号']}</td>
            <td class="text-center font-semibold">{row['所在月周']}</td>
            <td class="text-center font-semibold text-dark">{row['销售']}</td>
            <td class="text-center"><span class="badge {type_cls}">{row['兵种']}</span></td>
            <td class="font-semibold text-dark">{row['客户名称']}</td>
            <td>{row['工作类型']}</td>
            <td class="text-center"><span class="badge {support_cls}">{row['是否支持']}</span></td>
            <td>{row['支持人员']}</td>
            <td class="text-center text-muted">{row['计划时间']}</td>
            <td class="text-right font-mono font-semibold">{row['预计工时（h）']}</td>
            <td class="text-muted">{row['备注']}</td>
        </tr>
        """

    html = f"""
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="utf-8">
        <style>
            {CSS_VARIABLES}
            
            body {{
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
                background-color: var(--bg-main);
                color: var(--text-main);
                margin: 0;
                padding: 30px;
                width: 1200px;
                box-sizing: border-box;
            }}
            
            .container {{
                background-color: var(--bg-card);
                border-radius: 16px;
                box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -4px rgba(0, 0, 0, 0.05);
                border: 1px solid var(--border-light);
                padding: 30px;
            }}
            
            /* 标题区域 */
            .header {{
                border-bottom: 2px solid var(--accent-light);
                padding-bottom: 20px;
                margin-bottom: 24px;
                display: flex;
                justify-content: space-between;
                align-items: flex-end;
            }}
            
            .header-title h1 {{
                font-size: 26px;
                color: var(--text-dark);
                margin: 0 0 8px 0;
                font-weight: 700;
                letter-spacing: -0.5px;
            }}
            
            .header-title p {{
                font-size: 14px;
                color: var(--text-muted);
                margin: 0;
            }}
            
            .header-logo {{
                font-size: 13px;
                background: linear-gradient(135deg, var(--accent), #1d4ed8);
                color: white;
                padding: 6px 12px;
                border-radius: 20px;
                font-weight: 600;
            }}
            
            /* 数据面板 */
            .dashboard {{
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 20px;
                margin-bottom: 30px;
            }}
            
            .card {{
                background: var(--bg-main);
                border: 1px solid var(--border-light);
                border-radius: 12px;
                padding: 16px 20px;
                box-sizing: border-box;
            }}
            
            .card-label {{
                font-size: 13px;
                color: var(--text-muted);
                margin-bottom: 6px;
                font-weight: 500;
            }}
            
            .card-value {{
                font-size: 24px;
                font-weight: 700;
                color: var(--text-dark);
            }}
            
            .card-value span {{
                font-size: 14px;
                font-weight: 500;
                color: var(--text-muted);
                margin-left: 4px;
            }}
            
            /* 表格设计 */
            table {{
                width: 100%;
                border-collapse: separate;
                border-spacing: 0;
                margin-top: 10px;
                border-radius: 8px;
                overflow: hidden;
                border: 1px solid var(--border-light);
            }}
            
            th {{
                background-color: var(--primary-light);
                color: white;
                font-weight: 600;
                font-size: 14px;
                padding: 14px 12px;
                text-align: left;
                border-bottom: 1px solid var(--border-color);
            }}
            
            td {{
                padding: 12px;
                font-size: 13.5px;
                border-bottom: 1px solid var(--border-light);
                color: var(--text-main);
                word-break: break-all;
                line-height: 1.5;
            }}
            
            tr:nth-child(even) td {{
                background-color: var(--bg-main);
            }}
            
            tr:last-child td {{
                border-bottom: none;
            }}
            
            /* 辅助对齐类 */
            .text-center {{ text-align: center; }}
            .text-right {{ text-align: right; }}
            .font-semibold {{ font-weight: 600; }}
            .font-mono {{ font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; }}
            .text-dark {{ color: var(--text-dark); }}
            .text-muted {{ color: var(--text-muted); }}
            
            /* 徽章 Badge */
            .badge {{
                display: inline-block;
                padding: 3px 8px;
                border-radius: 6px;
                font-size: 12px;
                font-weight: 600;
            }}
            .badge-yes {{ background-color: var(--badge-yes-bg); color: var(--badge-yes-text); }}
            .badge-no {{ background-color: var(--badge-no-bg); color: var(--badge-no-text); }}
            .badge-af {{ background-color: var(--badge-af-bg); color: var(--badge-af-text); }}
            .badge-xf {{ background-color: var(--badge-xf-bg); color: var(--badge-xf-text); }}
            .badge-jf {{ background-color: var(--badge-jf-bg); color: var(--badge-jf-text); }}
            .badge-other {{ background-color: var(--badge-other-bg); color: var(--badge-other-text); }}
            
            /* 页脚 */
            .footer {{
                margin-top: 30px;
                text-align: center;
                font-size: 12px;
                color: var(--text-muted);
                border-top: 1px solid var(--border-light);
                padding-top: 15px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="dashboard">
                <div class="card">
                    <div class="card-label">计划总数</div>
                    <div class="card-value">{stats['total_count']}<span>项</span></div>
                </div>
                <div class="card">
                    <div class="card-label">预计总工时</div>
                    <div class="card-value">{stats['total_hours']}<span>小时</span></div>
                </div>
                <div class="card">
                    <div class="card-label">涉及客户数</div>
                    <div class="card-value">{stats['total_customers']}<span>家</span></div>
                </div>
                <div class="card">
                    <div class="card-label">销售团队</div>
                    <div class="card-value">{stats['total_sales']}<span>人</span></div>
                </div>
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th class="text-center" style="width: 50px;">序号</th>
                        <th class="text-center" style="width: 90px;">所在月周</th>
                        <th class="text-center" style="width: 80px;">销售</th>
                        <th class="text-center" style="width: 70px;">兵种</th>
                        <th style="width: 160px;">客户名称</th>
                        <th style="width: 140px;">工作类型</th>
                        <th class="text-center" style="width: 80px;">是否支持</th>
                        <th style="width: 100px;">支持人员</th>
                        <th class="text-center" style="width: 110px;">计划时间</th>
                        <th class="text-right" style="width: 80px;">工时(h)</th>
                        <th>备注</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>
            
            <div class="footer">
                生成时间：{stats['generation_time']} | 页面宽度已适配电脑端大图查看
            </div>
        </div>
    </body>
    </html>
    """
    return html

def generate_card_html(df, stats, week_info):
    """
    生成针对手机屏幕（竖屏）的精美卡片流 HTML 字符串，贯彻 taste-skill 殿堂级通透设计标准
    """
    # 增加空数据防御
    if df.empty:
        return f"""
        <!DOCTYPE html>
        <html lang="zh-CN">
        <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                {CSS_VARIABLES}
                body {{
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
                    background-color: var(--bg-main);
                    color: var(--text-main);
                    margin: 0;
                    padding: 75px 24px 24px 24px;
                    width: 100%;
                    box-sizing: border-box;
                }}
                .container {{
                    background-color: var(--bg-card);
                    border-radius: var(--radius-lg);
                    box-shadow: var(--shadow-stats);
                    border: 1px solid var(--border-light);
                    padding: 60px 24px;
                    text-align: center;
                }}
                .empty-icon {{
                    width: 64px;
                    height: 64px;
                    color: var(--text-muted);
                    margin: 0 auto 16px auto;
                }}
                .empty-title {{
                    font-size: 18px;
                    font-weight: 700;
                    color: var(--text-dark);
                    margin-bottom: 8px;
                }}
                .empty-desc {{
                    font-size: 13px;
                    color: var(--text-muted);
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <svg class="empty-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                    <circle cx="12" cy="12" r="10"></circle>
                    <line x1="12" y1="8" x2="12" y2="12"></line>
                    <line x1="12" y1="16" x2="12.01" y2="16"></line>
                </svg>
                <div class="empty-title">暂无本周技术安排</div>
                <div class="empty-desc">导入的数据表格中未发现有效计划记录。</div>
                <!-- 智能裁剪定位标记线 -->
                <div style="height: 2px; background-color: #0f172a; margin-top: 40px;"></div>
            </div>
        </body>
        </html>
        """

    # 根据销售进行分组，使用 sort=False 确保拼音字母排序不被覆盖
    grouped = df.groupby('销售', sort=False)
    
    sections_html = ""
    for sales_name, group in grouped:
        sales_hours = 0
        try:
            sales_hours = pd.to_numeric(group['预计工时（h）'], errors='coerce').fillna(0).sum()
        except Exception:
            pass
            
        group_cards_html = ""
        for _, row in group.iterrows():
            type_cls = get_badge_class(row['兵种'])
            
            # 协作状态深度融合：合并是否支持与支持人员，展示精致的线稿头像徽章，并增加边框防褪色
            is_support = str(row['是否支持']).strip() == "是" and str(row['支持人员']).strip()
            if is_support:
                collab_html = f"""
                <span class="badge badge-yes">
                    <svg class="badge-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
                        <circle cx="12" cy="7" r="4" />
                    </svg>
                    {row['支持人员']} 支持
                </span>
                """
            else:
                collab_html = '<span class="badge badge-no">否</span>'
            
            # 备注处理，如果为空则不显示备注区块，增加精致线稿便签小图标
            remark_html = ""
            if str(row['备注']).strip():
                remark_html = f"""
                <div class="remark-box">
                    <span class="field-label">
                        <svg class="badge-icon" style="vertical-align: -1px; margin-right: 4px;" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
                        </svg>备注
                    </span>
                    <span class="field-value text-muted">{row['备注']}</span>
                </div>
                """
            
            # 增加计划卡片的不对称聚焦律动（若为协作型项目，升级为 plan-card-accent 高亮聚焦卡）
            card_cls = "plan-card plan-card-accent" if is_support else "plan-card"
            
            group_cards_html += f"""
            <div class="{card_cls}">
                <!-- 殿堂级艺术感背景超大字号序号水印，绝对定位在右下角，立体穿插 -->
                <div class="card-watermark">{row['序号']}</div>
                
                <div class="card-header">
                    <!-- 左侧添加微渐变 Focus Anchor（视觉落脚蓝色聚焦条） -->
                    <div class="customer-name-wrapper">
                        <span class="customer-anchor"></span>
                        <span class="customer-name">{row['客户名称']}</span>
                    </div>
                    <div class="card-header-badges">
                        <span class="badge {type_cls}">{row['兵种']}</span>
                        <span class="badge badge-hours">{row['预计工时（h）']}h</span>
                    </div>
                </div>
                
                <div class="card-body">
                    <div class="card-grid">
                        <div class="grid-item">
                            <span class="field-label">工作类型</span>
                            <span class="field-value font-semibold text-dark">{row['工作类型']}</span>
                        </div>
                        <div class="grid-item">
                            <span class="field-label">计划时间</span>
                            <span class="field-value text-dark">{row['计划时间']}</span>
                        </div>
                    </div>
                    
                    <div class="card-grid border-top">
                        <div class="grid-item">
                            <span class="field-label">协作安排</span>
                            <span class="field-value">{collab_html}</span>
                        </div>
                    </div>
                    
                    {remark_html}
                </div>
            </div>
            """
            
        sections_html += f"""
        <div class="sales-section">
            <div class="sales-section-header">
                <div class="sales-info">
                    <span class="sales-avatar">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
                            <circle cx="12" cy="7" r="4" />
                        </svg>
                    </span>
                    <span class="sales-title">{sales_name}</span>
                </div>
                <div class="sales-meta">
                    <span class="meta-badge">{len(group)} 项计划</span>
                    <span class="meta-badge hours-badge">{sales_hours:g}h 工时</span>
                </div>
            </div>
            <div class="sales-cards">
                {group_cards_html}
            </div>
        </div>
        """

    html = f"""
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            {CSS_VARIABLES}
            
            body {{
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
                background-color: var(--bg-main);
                /* 引入双重径向弥散极光渐变，叠合 40px 手稿网格底纹，呈现极致通透且高品质的画报质感 */
                background-image: 
                    radial-gradient(circle at 10% 20%, rgba(37, 99, 235, 0.02) 0%, transparent 40%),
                    radial-gradient(circle at 90% 80%, rgba(234, 88, 12, 0.015) 0%, transparent 40%),
                    linear-gradient(rgba(15, 23, 42, 0.012) 1px, transparent 1px), 
                    linear-gradient(90deg, rgba(15, 23, 42, 0.012) 1px, transparent 1px);
                background-size: 100% 100%, 100% 100%, 40px 40px, 40px 40px;
                color: var(--text-main);
                margin: 0;
                padding: 75px 24px 30px 24px; /* 左右间距统一为 24px，实现完美左右对称，顶部预留 75px 安全防线 */
                width: 100%; /* 自动撑满 650px 宽度视口 */
                box-sizing: border-box;
            }}
            
            /* 手机专属头部（去 Emoji 改为精致 SVG 图标，顶部引入物理高光与深空极光渐变，配合精密径向网格） */
            .header {{
                position: relative;
                background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #020617 100%); /* 深空三色极光 */
                color: white;
                border-radius: var(--radius-lg);
                padding: 26px 20px;
                margin-bottom: 20px;
                display: flex;
                flex-direction: column;
                align-items: center;
                gap: 10px;
                border-top: 1px solid rgba(255, 255, 255, 0.18); /* 顶部微光折射 */
                box-shadow: 0 8px 20px rgba(37, 99, 235, 0.1), 0 2px 6px rgba(15, 23, 42, 0.15); /* 底部微偏色发光阴影 */
                overflow: hidden;
            }}
            
            /* 科技微网格背景层 */
            .header::before {{
                content: "";
                position: absolute;
                inset: 0;
                background-image: radial-gradient(rgba(255, 255, 255, 0.08) 1px, transparent 0);
                background-size: 14px 14px;
                pointer-events: none;
                z-index: 1;
            }}
            
            .header-decor, .header h1 {{
                position: relative;
                z-index: 2;
            }}
            
            .header-decor {{
                display: inline-flex;
                align-items: center;
                justify-content: center;
                width: 38px;
                height: 38px;
                background-color: rgba(255, 255, 255, 0.12);
                border-radius: 50%;
                color: #93c5fd; /* 莫兰迪蓝色高亮 */
            }}
            
            .header-icon {{
                width: 20px;
                height: 20px;
            }}
            
            .header h1 {{
                font-size: 20px;
                margin: 0;
                font-weight: 800;
                letter-spacing: 0.08em; /* 增大字距展现高级质感 */
                text-transform: uppercase;
                text-align: center;
                text-shadow: 0 2px 4px rgba(0, 0, 0, 0.15);
            }}
            
            /* 手机统计栏 Double-Bezel 双嵌套结构，集成微型 inline SVG 图标与独立数据信息块并列，打造成高端 Dashboard 组件 */
            .stats-bar-shell {{
                background-color: rgba(15, 23, 42, 0.04); /* 外层底座色 */
                border-radius: var(--radius-lg);
                padding: 4px;
                margin-bottom: 24px;
                border: 1px solid rgba(15, 23, 42, 0.05);
            }}
            
            .stats-bar {{
                display: flex;
                justify-content: space-between;
                background-color: var(--bg-card); /* 内部核心纯白卡 */
                border-radius: calc(var(--radius-lg) - 4px);
                padding: 14px 16px;
                box-shadow: inset 0 1px 1px rgba(255, 255, 255, 0.8), var(--shadow-stats);
            }}
            
            .stat-item {{
                position: relative;
                flex: 1;
                display: flex;
                align-items: center;
                gap: 10px;
                padding: 2px 0 6px 8px;
            }}
            
            .stat-item:not(:last-child) {{
                border-right: 1px solid var(--border-light);
            }}
            
            /* 底部精密 2px 莫兰迪彩色指示线条，极大强化仪表盘设计感 */
            .stat-item::after {{
                content: "";
                position: absolute;
                bottom: 0;
                left: 12px;
                right: 12px;
                height: 2px;
                border-radius: 1px;
            }}
            .stat-item-1::after {{ background-color: var(--accent); }}
            .stat-item-2::after {{ background-color: #ea580c; }}
            .stat-item-3::after {{ background-color: #6366f1; }}
            
            .stat-icon-wrapper {{
                display: inline-flex;
                align-items: center;
                justify-content: center;
                width: 32px;
                height: 32px;
                border-radius: 8px;
                background-color: var(--bg-main);
                flex-shrink: 0;
            }}
            
            .text-orange {{ color: #ea580c; background-color: #fff7ed; }}
            .text-purple {{ color: #6366f1; background-color: #eef2ff; }}
            .text-accent {{ color: var(--accent); background-color: var(--accent-light); }}
            
            .stat-icon {{
                width: 16px;
                height: 16px;
            }}
            
            .stat-info-block {{
                display: flex;
                flex-direction: column;
                gap: 2px;
            }}
            
            .stat-label {{
                font-size: 10px;
                color: var(--text-muted);
                text-transform: uppercase;
                font-weight: 700;
                letter-spacing: 0.06em;
            }}
            
            .stat-val {{
                font-size: 18px;
                font-weight: 800;
                color: var(--text-dark);
                font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
                line-height: 1.1;
            }}
            
            .stat-val span {{
                font-size: 10px;
                font-weight: 600;
                color: var(--text-muted);
                margin-left: 2px;
            }}
            
            /* 分组模块 */
            .sales-section {{
                margin-bottom: 24px;
            }}
            
            .sales-section-header {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 4px;
                margin-bottom: 12px;
                position: relative;
            }}
            
            /* 细密渐变虚线分栏引线，连接姓名与右侧指标，提升精密科技质感 */
            .sales-section-header::after {{
                content: "";
                flex-grow: 1;
                border-bottom: 1px dashed rgba(15, 23, 42, 0.15);
                margin: 0 16px;
                height: 0;
            }}
            
            .sales-info {{
                display: flex;
                align-items: center;
            }}
            
            /* 分组头部 Avatar 质感提升 */
            .sales-avatar {{
                display: inline-flex;
                align-items: center;
                justify-content: center;
                width: 24px;
                height: 24px;
                background-color: var(--accent-light);
                border-radius: 50%;
                color: var(--accent); /* 高质感蓝色线稿头像圆圈 */
                margin-right: 8px;
            }}
            
            .sales-avatar svg {{
                width: 14px;
                height: 14px;
            }}
            
            .sales-title {{
                font-size: 16px;
                font-weight: 700;
                color: var(--text-dark);
            }}
            
            .sales-meta {{
                display: flex;
                gap: 6px;
                order: 3; /* 将指标块排在线条右侧 */
            }}
            
            .meta-badge {{
                font-size: 12px;
                background-color: #eff6ff; /* 浅蓝底 */
                color: var(--accent);      /* 亮蓝字 */
                padding: 4px 10px;
                border-radius: 12px;
                font-weight: 700;          /* 极粗字体 */
                border: 1px solid rgba(37, 99, 235, 0.12);
                box-shadow: 0 1px 2px rgba(15, 23, 42, 0.02);
            }}
            
            .hours-badge {{
                background-color: #fff7ed; /* 浅橙底 */
                color: #ea580c;            /* 橙色字 */
                border-color: rgba(234, 88, 12, 0.12);
            }}
            
            /* 卡片定义 - 扁平悬浮，解决 box-in-box 嵌套感 */
            .sales-cards {{
                display: flex;
                flex-direction: column;
                gap: 12px;
            }}
            
            .plan-card {{
                position: relative;
                overflow: hidden;
                background-color: var(--bg-card);
                border: 1px solid rgba(15, 23, 42, 0.08); /* 极致细致的外边框线 */
                border-radius: var(--radius-lg);
                padding: 16px;
                box-shadow: var(--shadow-card); /* 双重浮雕微阴影 */
            }}
            
            /* 重点协作项目 Accent Card（高亮聚焦卡），打破上下纯长条的单调性，制造横向视线律动 */
            .plan-card-accent {{
                background: linear-gradient(180deg, #ffffff 0%, #fcfdfe 100%);
                border: 1px solid rgba(37, 99, 235, 0.18);
                box-shadow: 0 4px 20px rgba(37, 99, 235, 0.04), 0 2px 4px rgba(15, 23, 42, 0.01), inset 0 1px 0 rgba(255, 255, 255, 0.95);
            }}
            
            /* 高档艺术感背景超大字号序号水印 */
            .card-watermark {{
                position: absolute;
                bottom: 8px;
                right: 14px;
                font-size: 72px; /* 进一步放大 */
                font-weight: 900;
                color: rgba(15, 23, 42, 0.022); /* 水印颜色调浅，表达极致清透感 */
                font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
                line-height: 1;
                pointer-events: none;
                z-index: 1;
            }}
            
            .card-header {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 12px;
                position: relative;
                z-index: 2;
            }}
            
            .customer-name-wrapper {{
                display: flex;
                align-items: center;
                gap: 8px;
                max-width: 70%;
            }}
            
            /* 视觉聚焦锚点（微渐变蓝色竖条且带有圆角发光投影，极具高档质感） */
            .customer-anchor {{
                display: inline-block;
                width: 3.5px;
                height: 14px;
                background: linear-gradient(to bottom, var(--accent), #60a5fa);
                border-radius: 2px;
                flex-shrink: 0;
                box-shadow: 0 1px 3px rgba(37, 99, 235, 0.2);
            }}
            
            .plan-card-accent .customer-anchor {{
                width: 5px; /* 高亮聚焦卡竖线加粗 */
                background: linear-gradient(to bottom, #2563eb, #3b82f6);
                box-shadow: 0 1px 6px rgba(37, 99, 235, 0.4); /* 发光增强 */
            }}
            
            .card-header-badges {{
                display: flex;
                gap: 6px;
                align-items: center;
            }}
            
            .customer-name {{
                font-size: 15px;
                font-weight: 800;
                color: var(--text-dark);
                line-height: 1.35;
            }}
            
            .card-body {{
                display: flex;
                flex-direction: column;
                gap: 10px;
                position: relative;
                z-index: 2;
            }}
            
            .card-grid {{
                display: flex;
                justify-content: space-between;
                gap: 10px;
            }}
            
            .grid-item {{
                flex: 1;
                display: flex;
                flex-direction: column;
                gap: 2px;
            }}
            
            .field-label {{
                font-size: 10px;
                color: var(--text-muted);
                font-weight: 700;
                letter-spacing: 0.05em;
                text-transform: uppercase;
                margin-bottom: 2px;
            }}
            
            .field-value {{
                font-size: 13.5px;
                color: var(--text-main);
                line-height: 1.45;
            }}
            
            .border-top {{
                border-top: 1px solid rgba(15, 23, 42, 0.04); /* 用超细浅分界实线，增加精细度 */
                padding-top: 8px;
                margin-top: 2px;
            }}
            
            /* 优化备注框的左侧边框颜色与背景透光感，改为精致的微蓝透光背景 */
            .remark-box {{
                background: rgba(37, 99, 235, 0.03); /* 高级淡蓝透光纸张感 */
                padding: 10px 12px;
                border-radius: var(--radius-sm);
                border-left: 3px solid #60a5fa; 
                margin-top: 4px;
                font-size: 12.5px;
                line-height: 1.6;
            }}
            
            /* 徽章 */
            .badge {{
                display: inline-flex;
                align-items: center;
                padding: 3px 7px;
                border-radius: var(--radius-sm);
                font-size: 11px;
                font-weight: 600;
                width: fit-content;
                line-height: 1;
            }}
            .badge-yes {{ 
                background-color: var(--badge-yes-bg); 
                color: var(--badge-yes-text); 
                border: 1px solid rgba(22, 101, 52, 0.12);
            }}
            .badge-no {{ 
                background-color: var(--badge-no-bg); 
                color: var(--badge-no-text); 
                border: 1px solid rgba(71, 85, 105, 0.1);
            }}
            .badge-af {{ 
                background-color: var(--badge-af-bg); 
                color: var(--badge-af-text); 
                border: 1px solid rgba(3, 105, 161, 0.12);
            }}
            .badge-xf {{ 
                background-color: var(--badge-xf-bg); 
                color: var(--badge-xf-text); 
                border: 1px solid rgba(13, 148, 136, 0.12);
            }}
            .badge-jf {{ 
                background-color: var(--badge-jf-bg); 
                color: var(--badge-jf-text); 
                border: 1px solid rgba(79, 70, 229, 0.12);
            }}
            .badge-other {{ 
                background-color: var(--badge-other-bg); 
                color: var(--badge-other-text); 
                border: 1px solid rgba(100, 116, 139, 0.1);
            }}
            
            /* 工时 Badge 胶囊化 */
            .badge-hours {{
                background-color: #fff7ed;
                color: #ea580c;
                border: 1px solid rgba(234, 88, 12, 0.12);
                border-radius: 12px; /* 胶囊化 */
                font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
            }}
            
            .badge-icon {{
                width: 12px;
                height: 12px;
                margin-right: 4px;
                display: inline-block;
                color: inherit;
            }}
            
            .text-accent {{ color: var(--accent); }}
            .font-bold {{ font-weight: 700; }}
            .font-mono {{ font-family: monospace; }}
            
            .footer {{
                text-align: center;
                margin-top: 30px;
                font-size: 11px;
                color: var(--text-muted);
                border-top: 1px solid var(--border-light);
                padding-top: 15px;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <div class="header-decor">
                <svg class="header-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                    <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
                    <line x1="16" y1="2" x2="16" y2="6"></line>
                    <line x1="8" y1="2" x2="8" y2="6"></line>
                    <line x1="3" y1="10" x2="21" y2="10"></line>
                </svg>
            </div>
            <h1>江西办 {week_info} 技术安排</h1>
        </div>
        
        <div class="stats-bar-shell">
            <div class="stats-bar">
                <div class="stat-item stat-item-1">
                    <span class="stat-icon-wrapper text-accent">
                        <svg class="stat-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                            <polyline points="9 11 12 14 22 4"></polyline>
                            <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"></path>
                        </svg>
                    </span>
                    <div class="stat-info-block">
                        <div class="stat-label">计划数</div>
                        <div class="stat-val">{stats['total_count']}<span>项</span></div>
                    </div>
                </div>
                <div class="stat-item stat-item-2">
                    <span class="stat-icon-wrapper text-orange">
                        <svg class="stat-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                            <circle cx="12" cy="12" r="10"></circle>
                            <polyline points="12 6 12 12 16 14"></polyline>
                        </svg>
                    </span>
                    <div class="stat-info-block">
                        <div class="stat-label">预计工时</div>
                        <div class="stat-val">{stats['total_hours']}<span>h</span></div>
                    </div>
                </div>
                <div class="stat-item stat-item-3">
                    <span class="stat-icon-wrapper text-purple">
                        <svg class="stat-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                            <circle cx="9" cy="7" r="4"></circle>
                            <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                            <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                        </svg>
                    </span>
                    <div class="stat-info-block">
                        <div class="stat-label">销售人数</div>
                        <div class="stat-val">{stats['total_sales']}<span>人</span></div>
                    </div>
                </div>
            </div>
        </div>
        
        {sections_html}
        
        <div class="footer">
            生成时间：{stats['generation_time']}
        </div>
        <!-- 智能裁剪定位标记线 -->
        <div style="height: 2px; background-color: #0f172a; margin-top: 30px;"></div>
    </body>
    </html>
    """
    return html

def calculate_stats(df):
    """
    计算数据集统计指标，鲁棒求和工时
    """
    stats = {}
    stats['total_count'] = len(df)
    
    # 算工时，使用 pd.to_numeric 代替手写转换
    total_hours = 0
    try:
        total_hours = pd.to_numeric(df['预计工时（h）'], errors='coerce').fillna(0).sum()
    except Exception:
        pass
    stats['total_hours'] = f"{total_hours:g}"
    
    # 算客户数
    stats['total_customers'] = df['客户名称'].nunique()
    # 算销售数
    stats['total_sales'] = df['销售'].nunique()
    
    stats['generation_time'] = datetime.datetime.now().strftime("%Y-%m-%d")
    return stats

def main():
    args = parse_args()
    
    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        print(f"错误：找不到文件 {input_path}")
        sys.exit(1)
        
    output_dir = args.output_dir
    if not output_dir:
        output_dir = os.path.dirname(input_path)
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    
    # 1. 数据清理与加工
    print("正在加载和处理数据...")
    df = clean_and_prepare_data(input_path)
    
    # 提取所在月周作为顶部标题
    week_info = ""
    if '所在月周' in df.columns:
        non_empty_weeks = df[df['所在月周'].astype(str).str.strip() != '']['所在月周']
        if not non_empty_weeks.empty:
            week_info = str(non_empty_weeks.iloc[0]).strip()
    if not week_info:
        week_info = "本周"
    
    # 2. 导出整理后的 Excel
    excel_out = os.path.join(output_dir, f"{base_name}_整理.xlsx")
    print(f"正在导出整理后的 Excel：{excel_out}")
    df.to_excel(excel_out, index=False)
    beautify_excel(excel_out)
    
    # 3. 计算指标
    stats = calculate_stats(df)
    
    # 4. 生成临时 HTML 文件并截图
    temp_dir = os.path.join(output_dir, ".tmp_render")
    os.makedirs(temp_dir, exist_ok=True)
    
    card_html_path = os.path.join(temp_dir, "card_view.html")
    card_png_path = os.path.join(output_dir, f"{base_name}_手机卡片流.png")
    card_content = generate_card_html(df, stats, week_info)
    with open(card_html_path, "w", encoding="utf-8") as f:
        f.write(card_content)
        
    # 5. 调用 Headless Chrome 进行截图
    chrome_path = find_chrome_binary()
    if not chrome_path:
        print("错误：未能在系统中找到 Chrome/Chromium 浏览器。请安装 Chrome 或设置 CHROME_PATH 环境变量。")
        sys.exit(1)
    
    # 动态估算渲染高度，自适应调整视口，防截断并提升小图性能
    sales_count = df['销售'].nunique() if '销售' in df.columns else 0
    row_count = len(df)
    
    if df.empty:
        card_height = 800
    else:
        estimated_height = 450 + (sales_count * 80) + (row_count * 180) + 300
        card_height = max(1200, min(15000, estimated_height))
    
    print("正在调用 Chrome Headless 渲染图片...")
    print(f"正在渲染手机卡片流长图 (650x{card_height} 自动裁剪) -> {card_png_path}")
    cmd_card = [
        chrome_path,
        "--headless",
        "--disable-gpu",
        "--force-device-scale-factor=2",
        f"--screenshot={card_png_path}",
        f"--window-size=650,{card_height}",
        card_html_path
    ]
    subprocess.run(cmd_card, capture_output=True)
    
    # 执行智能高精度高度裁剪
    crop_image_by_marker(card_png_path)
    
    # 清理临时渲染 HTML
    try:
        os.remove(card_html_path)
        os.rmdir(temp_dir)
    except Exception:
        pass
        
    print("\n转换全部完成！已成功生成以下文件：")
    print(f"1. [整理后的 Excel] {excel_out}")
    print(f"2. [手机卡片流长图] {card_png_path}")

if __name__ == "__main__":
    main()
