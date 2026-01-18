from flask import Flask, render_template_string, request, jsonify
import pandas as pd
import numpy as np
import os

app = Flask(__name__)
# 允许上传的文件格式
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# 检查文件格式是否合法
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# 成绩分析核心函数（适配常见成绩表，可直接改）
def analyze_scores(file):
    df = pd.read_excel(file)
    # 提取成绩列（默认取最后一列，可手动指定列名如df['数学']）
    score_col = df.columns[-1]
    scores = df[score_col].dropna()
    
    # 核心统计
    total = len(scores)
    avg = round(np.mean(scores), 2)
    max_score = round(np.max(scores), 2)
    min_score = round(np.min(scores), 2)
    pass_num = len(scores[scores >= 60])
    pass_rate = round(pass_num / total * 100, 2)
    
    return {
        "科目": score_col,
        "参考人数": total,
        "平均分": avg,
        "最高分": max_score,
        "最低分": min_score,
        "及格人数": pass_num,
        "及格率": f"{pass_rate}%",
        "原始数据": df.head(10).to_html(index=False)  # 展示前10条数据
    }

# 内置HTML页面（不用单独写文件）
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>成绩统计分析工具</title>
    <style>
        body {max-width: 900px; margin: 20px auto; padding: 0 15px; font-family: Arial;}
        .upload-box {padding: 25px; border: 2px dashed #ccc; border-radius: 8px; margin-bottom: 20px;}
        button {background: #007bff; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer;}
        input[type=file] {margin: 10px 0;}
        .result-box {margin-top: 30px; padding: 20px; border: 1px solid #eee; border-radius: 8px;}
        table {width: 100%; border-collapse: collapse; margin: 15px 0;}
        th,td {border: 1px solid #ccc; padding: 8px; text-align: center;}
        th {background: #f5f5f5;}
    </style>
</head>
<body>
    <h1>成绩统计分析工具</h1>
    <div class="upload-box">
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xlsx,.xls" required>
            <br>
            <button type="submit">上传Excel并分析</button>
        </form>
    </div>

    {% if result %}
    <div class="result-box">
        <h3>分析结果</h3>
        <p>科目：{{result.科目}}</p>
        <p>参考人数：{{result.参考人数}} | 平均分：{{result.平均分}}</p>
        <p>最高分：{{result.最高分}} | 最低分：{{result.最低分}}</p>
        <p>及格人数：{{result.及格人数}} | 及格率：{{result.及格率}}</p>
        <h4>数据预览（前10条）</h4>
        {{result.原始数据|safe}}
    </div>
    {% endif %}
</body>
</html>
"""

# 首页+上传+分析 三合一接口
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 检查是否有文件上传
        if 'file' not in request.files:
            return render_template_string(HTML_TEMPLATE, result={"科目":"错误", "参考人数":0, "平均分":0, "最高分":0, "最低分":0, "及格人数":0, "及格率":"0%", "原始数据":"未选择文件"})
        file = request.files['file']
        if file.filename == '':
            return render_template_string(HTML_TEMPLATE, result={"科目":"错误", "参考人数":0, "平均分":0, "最高分":0, "最低分":0, "及格人数":0, "及格率":"0%", "原始数据":"文件名不能为空"})
        if file and allowed_file(file.filename):
            try:
                result = analyze_scores(file)
                return render_template_string(HTML_TEMPLATE, result=result)
            except Exception as e:
                return render_template_string(HTML_TEMPLATE, result={"科目":"分析失败", "参考人数":0, "平均分":0, "最高分":0, "最低分":0, "及格人数":0, "及格率":"0%", "原始数据":f"错误原因：{str(e)}"})
    # GET请求返回上传页面
    return render_template_string(HTML_TEMPLATE)

# 启动服务（适配Cloudflare部署，不用改）
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 8080)), debug=False)
