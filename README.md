# 上岸手账

个人考公学习网站，包含学习计划、专注计时、刷题统计、成语积累、错题导入、复习队列和自定义主题。

## 发布到 GitHub Pages

解压后，在本文件所在文件夹的地址栏输入 `powershell` 并回车，然后运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy.ps1
```

脚本会检查 GitHub CLI 和 Git，自动创建公开仓库、上传代码并启用 GitHub Pages。网站构建由 GitHub 自动完成，本机无需安装 Node.js。

如果希望手动操作，也可以逐行运行：

```powershell
npm install
git init
git add .
git commit -m "Create gongkao study site"
git branch -M main
gh repo create gongkao-study --public --source=. --remote=origin
$owner = gh api user --jq .login
gh api --method POST "repos/$owner/gongkao-study/pages" -f build_type=workflow
git push -u origin main
gh run watch
```

发布完成后的地址：

```text
https://你的GitHub用户名.github.io/gongkao-study/
```

如果 `gh api` 提示 Pages 已存在，可以忽略该提示，继续执行 `git push`。

## 数据说明

学习数据保存在当前浏览器的本地存储中，不会上传到 GitHub。更换浏览器、设备或清理浏览器数据前，请先导出重要内容。
