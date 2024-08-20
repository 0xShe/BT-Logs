import os
import re
import pandas as pd
from datetime import datetime

# 定义正则表达式来解析日志文件
log_pattern = re.compile(
    r'(?P<ip>[\d\.]+) - - \[(?P<datetime>[^\]]+)\] "(?P<method>GET|POST) (?P<request>[^"]+)" (?P<status>\d+) (?P<size>\d+) "-" "(?P<user_agent>[^"]+)"'
)

def parse_log_line(line):
    match = log_pattern.match(line)
    if match:
        return match.groupdict()
    return None

def convert_datetime(date_str):
    return datetime.strptime(date_str, '%d/%b/%Y:%H:%M:%S %z').strftime('%Y-%m-%d %H:%M:%S')

def process_log_files():
    log_files = [f for f in os.listdir('.') if f.endswith('.log')]
    
    for log_file in log_files:
        all_data = []
        with open(log_file, 'r') as file:
            for line in file:
                parsed_line = parse_log_line(line)
                if parsed_line:
                    parsed_line['datetime'] = convert_datetime(parsed_line['datetime'])
                    all_data.append(parsed_line)
        
        if all_data:
            df = pd.DataFrame(all_data)
            # 转换列名为中文
            df.columns = ['IP 地址', '日期时间', '请求方法', '请求地址', '状态码', '数据大小', '用户代理']
            excel_file_name = f"{os.path.splitext(log_file)[0]}.xlsx"
            df.to_excel(excel_file_name, index=False)
            
            # Analysis
            analysis_report(df, excel_file_name)

def analysis_report(df, excel_file_name):
    # 1. 提取出现次数最多的10个IP的详细信息
    top_ips = df['IP 地址'].value_counts().head(10)
    top_ips_details = []
    for ip in top_ips.index:
        ip_data = df[df['IP 地址'] == ip]
        ip_earliest = ip_data['日期时间'].min()
        ip_latest = ip_data['日期时间'].max()
        ip_monthly_counts = ip_data['日期时间'].apply(lambda x: datetime.strptime(x, "%Y-%m-%d %H:%M:%S").month).value_counts().reindex(range(1, 13), fill_value=0)
        ip_info = {
            'IP 地址': ip,
            '出现次数': top_ips[ip],
            '最早出现的时间': ip_earliest,
            '最后一次出现的时间': ip_latest,
            **{f'{month}月': ip_monthly_counts.get(month, 0) for month in range(1, 13)}
        }
        top_ips_details.append(ip_info)
    
    top_ips_df = pd.DataFrame(top_ips_details)

    # 2. 状态码统计
    all_status_codes = df['状态码'].unique()
    status_counts = df['状态码'].value_counts()
    status_summary = {f'状态码{code}': status_counts.get(str(code), 0) for code in sorted(all_status_codes.astype(int))}
    status_summary['其他状态码'] = status_counts.sum() - sum(status_summary.values())
    status_df = pd.DataFrame(status_summary, index=['次数']).T

    # 3. 被访问最多的前10个GET/POST地址及其访问情况
    top_requests = df['请求地址'].value_counts().head(10)
    request_details = []
    for request in top_requests.index:
        request_data = df[df['请求地址'] == request]
        request_ips = request_data['IP 地址'].value_counts().head(10)
        for ip, count in request_ips.items():
            request_info = {
                '请求地址': request,
                'IP 地址': ip,
                '访问次数': count
            }
            request_details.append(request_info)
    
    requests_df = pd.DataFrame(request_details)

    # 保存到Excel文件
    with pd.ExcelWriter(excel_file_name, mode='a', engine='openpyxl') as writer:
        top_ips_df.to_excel(writer, sheet_name='前10个IP详细信息', index=False)
        status_df.to_excel(writer, sheet_name='状态码统计', index=True)
        requests_df.to_excel(writer, sheet_name='最受欢迎的请求', index=False)

if __name__ == "__main__":
    process_log_files()
