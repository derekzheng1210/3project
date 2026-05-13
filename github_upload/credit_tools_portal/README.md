# 信用债研究平台部署包

这是可手工上传到 GitHub 的 Flask 部署目录。

## 本地运行

```powershell
pip install -r requirements.txt
$env:SITE_PASSWORD="your-password"
$env:SECRET_KEY="change-this-secret"
python app.py
```

访问 `http://127.0.0.1:5000`，先输入 `SITE_PASSWORD`。

## Render/云端建议

- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app`
- Environment Variables:
  - `SITE_PASSWORD`: 全站访问密码
  - `SECRET_KEY`: Flask session 密钥

## 数据更新

- 择券工具：后台上传 Excel。
- 策略仪表盘：本地运行 Wind 脚本生成 HTML 后上传。
- 利差监控：本地生成 `spread_data.js` 后上传；`spread_data.json` 可选上传留档。
- 曲线映射表固定随代码发布，位于 `config/映射表.xlsx`。