"""
生成ICS日历文件供Outlook订阅

使用方式：
    python generate_calendar.py
    
输出：
    events.ics （可通过GitHub raw链接订阅到Outlook）
"""

import re
from datetime import datetime
from zoneinfo import ZoneInfo

BOSTON_TZ = ZoneInfo("America/New_York")

def parse_time(time_str: str) -> tuple[str, str] | None:
    """解析时间字符串 '4:00 PM – 6:00 PM' -> ('16:00', '18:00')"""
    try:
        parts = time_str.split('–')
        if len(parts) != 2:
            return None
        
        start_str = parts[0].strip()
        end_str = parts[1].strip()
        
        # 解析起始时间
        start_match = re.match(r'(\d{1,2}):(\d{2})\s*(AM|PM)', start_str)
        if not start_match:
            return None
        start_hour = int(start_match.group(1))
        start_min = int(start_match.group(2))
        start_period = start_match.group(3)
        
        if start_period == 'PM' and start_hour != 12:
            start_hour += 12
        elif start_period == 'AM' and start_hour == 12:
            start_hour = 0
        
        # 解析结束时间
        end_match = re.match(r'(\d{1,2}):(\d{2})\s*(AM|PM)', end_str)
        if not end_match:
            return None
        end_hour = int(end_match.group(1))
        end_min = int(end_match.group(2))
        end_period = end_match.group(3)
        
        if end_period == 'PM' and end_hour != 12:
            end_hour += 12
        elif end_period == 'AM' and end_hour == 12:
            end_hour = 0
        
        return (f"{start_hour:02d}{start_min:02d}00", f"{end_hour:02d}{end_min:02d}00")
    except Exception:
        return None


def generate_ics(events_data: list) -> str:
    """
    生成ICS日历文件内容
    
    参数：
        events_data: 事件列表，每个事件包含：
            {
                'date': 'YYYYMMDD',
                'title': '事件标题',
                'time': '4:00 PM – 6:00 PM',
                'food': '食物描述',
                'location': '位置',
                'url': '活动链接',
                'source': '数据源缩写'
            }
    """
    
    now = datetime.now(BOSTON_TZ)
    ics_lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Harvard Food Events//harvard-food-events//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:Harvard Free Food Events",
        f"X-WR-CALDESC:Free food events at Harvard University",
        f"X-WR-TIMEZONE:America/New_York",
        "BEGIN:VTIMEZONE",
        "TZID:America/New_York",
        "BEGIN:STANDARD",
        "DTSTART:20231105T020000",
        "TZOFFSETFROM:-0400",
        "TZOFFSETTO:-0500",
        "END:STANDARD",
        "BEGIN:DAYLIGHT",
        "DTSTART:20240310T020000",
        "TZOFFSETFROM:-0500",
        "TZOFFSETTO:-0400",
        "END:DAYLIGHT",
        "END:VTIMEZONE",
    ]
    
    for event in events_data:
        try:
            date_str = event.get('date', '')
            time_str = event.get('time', '')
            
            # 解析时间
            time_result = parse_time(time_str)
            if not time_result:
                continue
            
            start_time, end_time = time_result
            dtstart = f"{date_str}T{start_time}"
            dtend = f"{date_str}T{end_time}"
            
            # 生成事件
            uid = f"{date_str}-{start_time}@harvard-food-events"
            title = event.get('title', 'Free Food Event')
            location = event.get('location', '')
            food = event.get('food', '')
            source = event.get('source', '')
            url = event.get('url', '')
            
            # 构建描述
            description = f"Food: {food}\\n\\nSource: {source}"
            if url:
                description += f"\\n\\nMore info: {url}"
            
            ics_lines.extend([
                "BEGIN:VEVENT",
                f"UID:{uid}",
                f"DTSTAMP:{now.isoformat().replace('+', 'Z').split('.')[0]}Z",
                f"DTSTART;TZID=America/New_York:{dtstart}",
                f"DTEND;TZID=America/New_York:{dtend}",
                f"SUMMARY:{title}",
                f"LOCATION:{location}",
                f"DESCRIPTION:{description}",
                f"URL:{url}" if url else "",
                f"CATEGORIES:{source}",
                "STATUS:CONFIRMED",
                "END:VEVENT",
            ])
        except Exception as e:
            print(f"Warning: Failed to parse event {event}: {e}")
            continue
    
    ics_lines.append("END:VCALENDAR")
    
    return "\n".join([line for line in ics_lines if line])


def extract_events_from_readme(readme_path: str) -> list:
    """从README.md中提取事件数据"""
    events = []
    
    with open(readme_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 提取日期标题（如 "## Friday, April 3"）
    date_pattern = r'## (\w+), (\w+) (\d+)'
    time_pattern = r'\| ([\d:]+ [AP]M.*?[AP]M) \|'
    event_pattern = r'\| [\d:]+ [AP]M.*?[AP]M \| \[(.*?)\].*?\| (.*?) \| (.*?) \| (.*?) \|'
    
    # 提取所有行
    lines = content.split('\n')
    
    current_date = None
    current_month = None
    current_year = 2026  # 假设是2026年
    
    for i, line in enumerate(lines):
        # 检查日期行
        date_match = re.search(date_pattern, line)
        if date_match:
            month_name = date_match.group(2)
            day = int(date_match.group(3))
            
            # 转换月份名称到数字
            months = {
                'January': 1, 'February': 2, 'March': 3, 'April': 4,
                'May': 5, 'June': 6, 'July': 7, 'August': 8,
                'September': 9, 'October': 10, 'November': 11, 'December': 12
            }
            month_num = months.get(month_name)
            if month_num:
                current_month = month_num
                current_date = f"{current_year}{month_num:02d}{day:02d}"
        
        # 检查事件行
        event_match = re.search(event_pattern, line)
        if event_match and current_date:
            time_str = re.search(time_pattern, line)
            if time_str:
                title = event_match.group(1)
                food = event_match.group(2)
                location = event_match.group(3)
                source = event_match.group(4)
                url = re.search(r'\[(.*?)\]\((.*?)\)', event_match.group(1))
                
                events.append({
                    'date': current_date,
                    'time': time_str.group(1),
                    'title': title,
                    'food': food.strip(),
                    'location': location.strip(),
                    'source': source.strip(),
                    'url': url.group(2) if url else ''
                })
    
    return events


if __name__ == "__main__":
    try:
        # 从README.md提取事件
        events = extract_events_from_readme("README.md")
        
        if not events:
            print("Warning: No events found in README.md")
        else:
            print(f"Found {len(events)} events")
        
        # 生成ICS
        ics_content = generate_ics(events)
        
        # 保存到文件
        output_file = "events.ics"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(ics_content)
        
        print(f"✅ Calendar generated: {output_file}")
        print(f"\n📅 Outlook订阅链接：")
        print(f"https://raw.githubusercontent.com/thetaaaaa/CrimsonEats/main/events.ics")
        print(f"\n在Outlook中：")
        print(f"1. 右键点击日历列表 -> '从Internet添加日历'")
        print(f"2. 粘贴上面的链接")
        print(f"3. 点击'订阅'")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
