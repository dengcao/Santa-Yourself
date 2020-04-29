# 圣诞老人你来做（Santa Yourself）

圣诞老人你来做 - Santa Yourself ，个性圣诞卡,圣诞节活动,跳舞的圣诞老人。

### 系统说明：

本程序开发于2008年12月，现在开源出来，希望能帮到有需要的人。

支持本程序，请到Gitee和GitHub给我们点Star！

Gitee：https://gitee.com/caozha/Santa-Yourself

GitHub：https://github.com/cao-zha/Santa-Yourself

### 安装方法

直接将源代码上传到您的网站根目录。

第一步，找到根目录下的data/sadan-user#2008-12-24.mdb文件，修改为其他复杂的文件名。

第二步，打开conn.asp，找到这行代码：

db_path="data/sadan-user#2008-12-24.mdb"

然后把sadan-user#2008-12-24.mdb修改为您刚才更改的数据库文件名即可，一定要修改，免得网站被黑。

第三步，打开：你的网站地址/index.asp ，就可以正常使用了。

### 注意事项

1、网站服务器的IIS必须 **开启父路径** ，否则程序会执行错误，会提示“不允许的父路径”的错误。[点此查看开启方法](https://my.oschina.net/dengzhenhua/blog/3295146)

2、网站目录必须开启写入权限，否则会安装失败。

3、程序使用了“化境HTTP上传程序 Version 2.1”，如果杀毒软件误报，你担心的话可以改成其他ASP上传组件。


### 版权所有

开发：草札 www.caozha.com

鸣谢：品络 www.pinluo.com  &ensp;  穷店 www.qiongdian.com

