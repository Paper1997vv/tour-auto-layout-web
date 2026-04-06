# App Bundle Notes

- 正式应用包路径：`/Users/X/Documents/自动化转化/TourAutoLayout.app`
- 文稿快捷方式：Finder alias，放在 `~/Documents/TourAutoLayout.app`
- 打包命令：

```bash
cd /Users/X/Documents/自动化转化
./scripts/package_app.sh
```

- 脚本会执行：
  - `swift build`
  - 复制最新可执行文件进 `.app`
  - 写入 `Info.plist` 和 `PkgInfo`
  - 重新创建文稿别名
  - 直接打开应用
