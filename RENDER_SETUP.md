# Render 部署 Google Drive 集成配置

## 1. Google Cloud Console 配置（一次性）

1. 前往 https://console.cloud.google.com 创建或选择项目
2. 启用 API：**Google Drive API**
3. 创建 OAuth 2.0 凭据：
   - 类型：**Web 应用**
   - 名称：`pd-translate`
   - 授权的重定向 URI：
     - `http://localhost:3000/auth/google/callback`（本地开发）
     - `https://pd-translate.onrender.com/auth/google/callback`（生产）
4. 复制 **Client ID** 和 **Client Secret**
5. OAuth 同意屏幕 → 添加测试用户（你自己 + 校对同事的 Google 邮箱）

## 2. Render 环境变量

在 Render Dashboard → Environment 中添加：

| 变量名 | 值 | 说明 |
|--------|-----|------|
| `GOOGLE_CLIENT_ID` | `123456789-xxx.apps.googleusercontent.com` | Google OAuth 客户端 ID |
| `GOOGLE_CLIENT_SECRET` | `GOCSPX-xxxx` | Google OAuth 客户端密钥 |
| `SESSION_SECRET` | 随机 32 位字符串 | Express session 加密密钥 |
| `NODE_ENV` | `production` | 运行环境 |

## 3. 部署步骤

1. 将代码推送到 GitHub
2. Render 自动检测并重新部署
3. 部署完成后访问网站
4. 点「登录 Google」测试认证流程
5. 登录后点「从 Google Drive 打开」测试文件列表

## 4. 共享文件使用流程

1. **主编** 登录 Google → 上传或创建校对文件
2. **主编** 在 Google Drive 网页版，右键 `pd-translate` 文件夹 → 共享
3. 添加校对者邮箱 → 权限设为「编辑者」
4. **校对者** 登录工作台 → 点「从 Google Drive 打开」
5. 列表中会显示主编共享的文件 → 双击打开
6. 编辑内容自动保存回主编的 Google Drive

## 5. 回滚方案

```bash
git checkout backup-before-google-drive
git push origin main --force
```

Render 会自动重新部署备份版本。
