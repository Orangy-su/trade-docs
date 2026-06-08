# 公网部署指南（Railway，免费，5分钟搞定）

## 方法一：Railway（推荐，完全免费）

Railway 是最简单的云部署平台，免费额度完全够用。

### 第一步：注册 Railway
访问 https://railway.app，用 GitHub 账号登录（需要先注册 GitHub）。

### 第二步：安装 Railway CLI（命令行工具）
在命令窗口运行：
```
npm install -g @railway/cli
```

### 第三步：登录 Railway
```
railway login
```

### 第四步：部署应用
在「外贸单据系统」文件夹里运行：
```
railway init
railway up
```

### 第五步：获取公网网址
```
railway open
```
会自动打开浏览器，显示你的公网网址，格式类似：
`https://qingfeng-docs-production.up.railway.app`

**把这个网址发给同事，任何地方都能用。**

---

## 方法二：Render（备选，也免费）

访问 https://render.com，连接 GitHub 仓库，选 Web Service，
设置：
- Build Command: `pip install -r requirements.txt`  
- Start Command: `python app_v2.py`

---

## 注意事项

1. **主数据手册**：部署到云端后，上传的主数据手册存在服务器上，重新部署会清空，建议定期备份。
2. **免费套餐限制**：Railway 免费版每月 $5 额度，一般够用；Render 免费版15分钟无访问会休眠（首次访问慢30秒）。
3. **安全性**：如果担心数据安全，可以在 app_v2.py 里加一个简单密码验证，把这个需求告诉 Claude 帮你加上。
