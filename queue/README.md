# 即查队列

页面「查这家」会在临时页 inquiry.html 生成一个 GitHub 提交链接，点一次绿色 Commit 就入队。
文件名 queue/<公司>-<产品KEY>.txt，内容 `公司名 | 产品KEY`（PLEX/FTPC/TWX/WIND/DMX）。
每天 09:00 / 20:30 自动任务读取、深查、写进 data/inquiries.json，然后删除任务文件。
本文件扩展名是 .md，不会被当成任务。
