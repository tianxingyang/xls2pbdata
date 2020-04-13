# Xls2pbdata
## 内容列表
- [背景](#背景)
- [安装](#安装)
  - [Requirements](#Requirements)
- [使用](#使用)
## 背景
自己做了两款半的手游，感觉转配置这块问题还挺多的，由于策划水平参差不齐，配出来的配置有些莫名其妙的问题，有一些是在转配置的时候就能发现的，有些只有在测试中才能发现。有些大厂都有自己成熟的转配置工具，可以在转配置的时候就发现很多问题，方便策划进行修改，一些创业公司或者小的游戏工作室在这里可能要摔一些跤，要么是转配置的工具处理错误不够完善，没法很快的定位配置错误的原因，要么是要花一些不必要的时间在开发转配置的工具上。
这个仓库的目标是：
1. 为有转配置档需求的朋友提供一个尽可能成熟的工具，减少这部分的开发成本。
2. 提供一个扩展性尽可能高的配置标准，适配各种游戏的配置需求。
3. ~~提升一下自己水平~~
## 安装
### <span id="Requirements">Requirements</span>:
* [python3](https://www.python.org/)
* [protobuf](https://github.com/protocolbuffers/protobuf) >= 3.11.2
* xlrd
* PyQt5

安装几个库的命令如下
```
python -m pip install protobuf==3.11.2 xlrd PyQt5
```
## 使用
1. 使用如下的文件目录结构
```
├── data //生成配置放在这个目录下
│   ├── client
│   ├── public
│   └── server
├── res //proto文件放在这里
│   ├── client.proto
│   ├── make_proto.bat
│   ├── protoc.exe
│   ├── public.proto
│   ├── server.proto
├── table // xlsx文件放在这个目录下
│   └── 配置.xlsx
└── xls2pbdata // 本项目
```
2. proto文件结构如下
```
message Foo
{
    message M
    {
        int32 id = 1; // 本配置中的id，方便程序（服务器或者客户端）中查找
        repeated int32 bar_0 = 2; // 具体的内容
        string bar_1 = 3; // 具体的内容
        bool bar_2 = 4; // 具体的内容
    }

    repeated M items_list  = 1; // 用来保存配置中所有的数据
}
```
3. excel文件举例如下

![Annotation 2020-04-13 231140.png](https://i.loli.net/2020/04/13/DIEUnrS9fjoQd4O.png)

4. 执行main.py

（默认转所有配置，包括 public/client/server）

5. 点击 select 选择excel文件
6. 点击 convert 生成对应的二进制文件 
