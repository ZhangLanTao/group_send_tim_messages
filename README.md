# group_send_tim_messages文档
用pymouse&amp;pykeyboard直接操作Windows-TIM发送信息给指定好友们

## 运行环境：

python3

pywin32

pyhook

PyUserInput

注：pyhook需要手动安装，详见github相关页面
注：github上传了已经编译的pyhook，将其放在python/Lib/site-packages下，在pyhook文件夹中powershell（管理员），运行pip install . ，可安装（不保证可用。）

## 使用说明：

将需要群发的好友列表写在data/numbers.txt里，每人一行格式为 xxxxx:名字，#后为注释，special以下的好友不发。

将自己的tim界面左上角头像和搜索框一起截图至data/search.png

根据Windows显示缩放比例设置scale（例如1.25）

在Windows剪贴板复制需要发送的内容（ctrl+c）打开tim界面，运行程序，确认发送时将send_message_confirm 置True，发送消息

