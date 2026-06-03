# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

将 PPT 文件（.ppt / .pptx）转换为带语音解说的 MP4 视频的命令行工具。每页 PPT 渲染成图片，每页的**备注文字**经 TTS 合成语音，再用语音时长驱动图片切换合成最终视频。

## 仓库结构

Maven 模块位于嵌套子目录 `ppt2video/`（即 `ppt2video/ppt2video/pom.xml`），而非仓库根目录。所有 `mvn` 命令都要在该子目录下执行。

## 构建与运行

```bash
# 构建（在 ppt2video/ 子目录下执行）
cd ppt2video && mvn package

# 运行（唯一参数是 PPT 文件的服务器全路径）
java -jar PptToVideoTool.jar /full/path/to/file.pptx
# 输出固定写到输入文件同目录：<同名>-pptToVideo.mp4
```

- 无测试代码，因此没有测试命令。
- **注意**：`pom.xml` 只配置了 `maven-compiler-plugin`，没有 shade/assembly/jar 插件，因此 `mvn package` **不会**自动生成 README 里提到的、含依赖与 `Main-Class` 的可运行 `PptToVideoTool.jar`。`Main-Class: ppt.PptToVideoTool` 仅声明在 `src/main/resource/META-INF/MANIFEST.MF` 中。要打出可运行 jar 需自行补充打包插件，或手动用 POI classpath 运行。

## 外部运行时依赖（不在 pom.xml 中）

工具大量通过 `Runtime.exec` shell-out，运行环境必须预装并配置：

- **ffmpeg**（README 验证版本 3.4.2）：需在 `PATH` 中，用于获取音频时长、拼接音频、生成视频、合并音视频。
- **科大讯飞 TTS 可执行文件**：路径**硬编码**在 `constants.ConstantParam.TTSFILEPATH`（默认 `/broker-tts/bin/tts`），换环境时必须改这里。
- **仅 Linux**：代码写死 `/bin/sh -c` 并依赖 `find`/`rm`/`touch`/`echo`/`grep`/`cut`/`sed`。README 注明 Linux 已验证、Windows 未验证。

## 核心架构

入口 `ppt.PptToVideoTool.main`，整个流水线靠类的静态可变字段（`voiceMap`、`pptToVideoTempFilePath`、`fileDirPath`、`FileName`）串联状态，按以下顺序执行：

1. **渲染 + 抽备注**：按后缀分流——`.ppt` 走 `convertToImage2003`（POI HSLF / `poi-scratchpad`），`.pptx` 走 `convertToImage2007`（POI XSLF / `poi-ooxml`）。每页用 `Graphics2D` 以 **3 倍**尺寸渲染成 jpeg，同时把该页备注文字存入 `voiceMap`（key 为 `voice_N.wav`）。两版统一把字体设为「宋体」防乱码。
2. **TTS 合成**：`createVoice` → `getVoicePath` → `commandExecutor`，对每条备注调用外部 TTS 生成 `voice_N.wav`。进程超时由 `ConstantParam.PROCESSTIMEOUT`（10 秒）控制。
3. **音频拼接**：`combineMp3` 用 ffmpeg 逐个读出 wav 的 `Duration`（返回「文件名 → 时长」map），再 `concat` 成一份 `combined.wav`。
4. **合成视频**：`createVideo` 写出 ffmpeg concat demuxer 的 `temp.txt`（每张图配对应语音时长），先生成无声 `noVoice.mp4` → `noVoice.avi`，再把 `combined.wav` 混入，输出最终 `<FileName>-pptToVideo.mp4`。
5. **清理**：`deleteTempFile` 删除所有中间文件及临时目录。

## 关键约定与坑

- **没有备注的页会被整页跳过**：渲染图片与抽备注在同一循环里，备注为空则 `continue`，该页既不生成图片也不生成语音。所以最终视频只包含「有备注」的页，图片与语音是强耦合的一一对应关系。
- **备注文本预处理**：写入 `voiceMap` 前会做 `@time=` → `p`、`]` → `000]` 的替换（TTS 标记转换），改备注处理逻辑时注意保留。
- **临时目录推导**：把输入路径的扩展名替换成 `/` 作为中间文件目录；后缀用 `filePath.substring(filePath.indexOf('.'))`（首个点，非 `lastIndexOf`），若**目录名里含点**会解析错误。
- 渲染倍数硬编码为 `3`（`convertToImageXXXX(..., 3)`），决定输出分辨率。
- 错误处理多为 `e.printStackTrace()` 后继续，外部命令失败不一定中断流程，排查问题时优先看 stdout 打印的逐条命令日志。
