# ppt2video
ppt文件转换为MP4工具类

## 1 功能描述

将输入的ppt文件转化成视频，视频是每页ppt和ppt的备注文字转化成的语音合成

* 将每一页的ppt切成图片，每一页ppt备注文字转化成语音
* 将所有的语音合成一份完整的语音，以语音长度为视频长度，与图片合成最终视频
* 视频中每段备注文字语音对应每页ppt，每段语音结束视频页面跳转到下一页ppt



## 2 参考输入输出

* 输入：待转化ppt文件路径
* 输出：转化后视频文件路径

~~~java
// 例如：
// 输入 - /home/hsfstore/hsStoredata/data/00/00/wKgh_V4EZzaEUj9wAAAAAAAAAAA79.pptx
// 输出 - /home/hsfstore/hsStoredata/data/00/00/wKgh_V4EZzaEUj9wAAAAAAAAAAA79-pptToVideo.mp4
~~~



## 3 调用方式

```cmake
# ssh
java -jar PptToVideoTool.jar [参数]
```



## 4 参数说明

* 目前仅支持单个参数，参数类型为String，为ppt在服务器上的全路径
* 后续支持TTS运行目录

## 5 所需环境说明

* `TTS`: 科大讯飞`tts`包
* `FFmpeg`: `version-3.4.2`

## 6 运行配置

* `ConstantParam`类中调整`TTSFILEPATH`字段为实际安装目录
* 科大讯飞的`TTS`和`FFmpeg`目前都支持在windows上安装，测试时可以用windows版本
* 本工程在`Liunx`服务器上验证通过，`windows`尚未验证