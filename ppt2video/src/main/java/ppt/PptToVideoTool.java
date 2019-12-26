package ppt;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.*;

import static constants.ConstantParam.*;


/**
 * 文件描述
 * @ProjectName: PptToVideoTool.jar
 * @Author: ching7
 * @date: 2019-12-26 14:00
 * @Version: 1.0
 * @note PptToVideoTool.jar
 * 功能描述：
 * 将输入的ppt文件转化成视频
 *     1.将每一页的ppt切成图片，每一页ppt备注文字转化成语音
 *     2.将所有的语音合成一份完整的语音，以语音长度为视频长度，与图片合成最终视频
 *     3.视频中每段语音对应每页ppt，每段语音结束视频页面跳转到下一页ppt
 *
 * 输入：带转化ppt文件路径
 * 输出：转化后视频文件路径
 * 例如：
 * 输入 - /home/hsfstore/hsStoredata/data/00/00/wKgh_V4EZzaEUj9wAAAAAAAAAAA79.pptx
 * 输出 - /home/hsfstore/hsStoredata/data/00/00/wKgh_V4EZzaEUj9wAAAAAAAAAAA79-pptToVideo.mp4
 *
 * 调用方式：
 *     java -jar PptToVideoTool.jar [参数]
 *
 * 参数说明：
 *     目前仅支持单个参数，参数类型为String,为ppt在服务器上的全路径
 *
 * 所需环境说明：
 *     tts: 科大讯飞tts包
 *     ffmpeg: version-3.4.2
 **/
public class PptToVideoTool {

    /**
     * ppt转视频中间文件存储目录
     */
    private static String pptToVideoTempFilePath = "";

    /**
     * 当前文件存储目录
     */
    private static StringBuilder fileDirPath = new StringBuilder();

    /**
     * 当前文件名称
     */
    private static String FileName = "";

    /**
     * 待合并话术
     */
    private static Map<String, String> voiceMap = new LinkedHashMap<>();


    public static void main(String[] args) throws Exception {
        // 参数1：待转化文件全路径
        // filePath: /root/storage/XX/XX/xxx.ppt
        String filePath = args[0];
        System.out.println("输入参数：" + filePath);
        // ppt转化为图片地址list
        List<String> list = null;
        // 带转化ppt文件后缀
        String suffix = filePath.substring(filePath.indexOf('.'));
        // 带转化ppt文件名称（无后缀）
        FileName = filePath.split("/")[filePath.split("/").length - 1].replace(suffix, "");
        // 转化临时文件夹地址
        pptToVideoTempFilePath = filePath.replace(suffix, "/");
        String cmdMakir = "mkdir " + pptToVideoTempFilePath;
        exec(cmdMakir);
        String[] fastDfsStoragePathArr = filePath.split("/");
        for (int i = 0; i < fastDfsStoragePathArr.length; i++) {
            if (i != fastDfsStoragePathArr.length - 1) {
                fileDirPath.append("/" + fastDfsStoragePathArr[i]);
            }
        }
        fileDirPath.append("/");
        System.out.println("ppt转视频中间文件生成目录：" + cmdMakir + "生成目录结果：" + cmdMakir);
        System.out.println("ppt转视频中间文件存储目录:" + pptToVideoTempFilePath);
        System.out.println("当前文件存储目录:" + fileDirPath + "/");
        System.out.println("当前文件名称:" + FileName);
        // 不同格式的ppt不同处理
        if (PPT.equals(suffix)) {
            list = convertToImage2003(filePath, pptToVideoTempFilePath, 3);
        } else if (PPTX.equals(suffix)) {
            list = convertToImage2007(filePath, pptToVideoTempFilePath, 3);
        }
        if (!list.isEmpty()) {
            //判断是否有已转换完成的视频，删除上次的
            String cmdDeleteTmpCombinedMp4 = "find " + fileDirPath + " -name '" + fileDirPath + FileName + "-pptToVideo.mp4' " + " -exec rm -f {} \\;";
            exec(cmdDeleteTmpCombinedMp4);
            // 生成话术
            createVoice();
            // 合并话术
            Map<String, String> map = combineMp3(voiceMap);
            // 合并视频
            createVideo(filePath, list, suffix, map);
            //删除临时文件
            deleteTempFile();
        }
        System.out.println("输出：ppt转视频结果文件位置" + fileDirPath + FileName + "-pptToVideo.mp4");
    }

    /**
     * 将PPT文件转换成image
     *
     * @param originalFileName   //PPT文件路径 如：d:/demo/demo1.ppt
     * @param targetImageFileDir //转换后的图片保存路径 如：d:/demo/pptImg
     * @return 图片名列表
     * @deprecated * 图片转化的格式字符串 ，如："jpg"、"jpeg"、"bmp" "png" "gif" "tiff"
     */
    public static List<String> convertToImage2003(String originalFileName, String targetImageFileDir, int times) {
        // PPT转成图片后所有名称集合
        List<String> picNames = new ArrayList<>();
        FileInputStream originalFileInputStream = null;
        FileOutputStream originalFileOutStream = null;
        HSLFSlideShow oneHSLFSlideShow = null;
        try {
            try {
                originalFileInputStream = new FileInputStream(originalFileName);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            try {
                oneHSLFSlideShow = new HSLFSlideShow(originalFileInputStream);
                originalFileInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            // 获取PPT每页的大小（宽和高度）
            Dimension onePageSize = oneHSLFSlideShow.getPageSize();
            // 获得PPT文件中的所有的PPT页面（获得每一张幻灯片）,并转为一张张的播放片
            List<HSLFSlide> pptPageSlideList = oneHSLFSlideShow.getSlides();
            // 下面循环的主要功能是实现对PPT文件中的每一张幻灯片进行转换和操作
            //保存话术内容
            StringBuffer bf = null;
            // 共循环ppt页数的次数
            int length = pptPageSlideList.size();
            for (int i = 0; i < length && pptPageSlideList.get(i).getNotes() != null; i++) {
                // 这几个循环只要是设置字体为宋体，防止中文乱码。获取ppt备注
                List<List<HSLFTextParagraph>> oneTextParagraphs = pptPageSlideList.get(i).getNotes().getTextParagraphs();
                if (!oneTextParagraphs.isEmpty()) {
                    bf = new StringBuffer();
                    for (List<HSLFTextParagraph> list : oneTextParagraphs) {
                        int index = 0;
                        for (HSLFTextParagraph hslfTextParagraph : list) {
                            if (!"".equals(hslfTextParagraph + "")) {
                                List<HSLFTextRun> hslFTextRunList = hslfTextParagraph.getTextRuns();
                                for (int j = 0; j < hslFTextRunList.size(); j++) {
                                    // 如果PPT在WPS中保存过，hslFTextRunList.get(j).getFontSize();的值为0或者26040，
                                    // 因此首先识别当前文本框内的字体尺寸是否为0或者大于26040，则设置默认的字体尺寸。
                                    // 设置字体大小
                                    Double size = hslFTextRunList.get(j).getFontSize();
                                    if (size == null || (size <= 0) || (size >= 26040)) {
                                        hslFTextRunList.get(j).setFontSize(20.0);
                                    }
                                    // 设置字体样式为宋体
                                    hslFTextRunList.get(j).setFontFamily("宋体");
                                }
                                String text = list.get(index).toString();
                                bf.append(text);
                            }
                            index++;
                        }
                    }
                    if ("".equals(bf.toString())) {
                        continue;
                    }
                    // 创建BufferedImage对象，图像的尺寸为原来的每页的尺寸*倍数times
                    BufferedImage oneBufferedImage = new BufferedImage(onePageSize.width * times, onePageSize.height * times, BufferedImage.TYPE_INT_RGB);
                    Graphics2D oneGraphics2D = oneBufferedImage.createGraphics();
                    // 设置转换后的图片背景色为白色
                    oneGraphics2D.setPaint(Color.white);
                    // 将图片放大times倍
                    oneGraphics2D.scale(times, times);
                    oneGraphics2D.fill(new Rectangle2D.Float(0, 0, onePageSize.width * times, onePageSize.height * times));
                    pptPageSlideList.get(i).draw(oneGraphics2D);
                    // 设置图片的存放路径和图片格式，注意生成的图片路径为绝对路径，最终获得各个图像文件所对应的输出流对象
                    try {

                        // 话术名字
                        String str = bf.toString();
                        str = str.replaceAll("@time=", "p").replaceAll("]", "000]");
                        voiceMap.put("voice_" + (i + 1) + ".wav", str);

                        String imgName = targetImageFileDir + "img_" + (i + 1) + ".jpeg";
                        File jpegFile = new File(imgName);
                        File dir = jpegFile.getParentFile();
                        if (!dir.exists()) {
                            dir.mkdirs();
                        }

                        //如果图片存在，则不再生成
//                        if (jpegFile.exists()) {
//                            continue;
//                        }

                        picNames.add(imgName);
                        originalFileOutStream = new FileOutputStream(imgName);
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    }
                    // 转换后的图片文件保存的指定的目录中
                    try {
                        ImageIO.write(oneBufferedImage, "jpeg", originalFileOutStream);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        } finally {
            try {
                if (originalFileOutStream != null) {
                    originalFileOutStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return picNames;
    }

    public static List<String> convertToImage2007(String filePath, String imgFile, int times) {
        List<String> picNames = new ArrayList<>();
        FileInputStream is = null;
        try {
            is = new FileInputStream(filePath);
            XMLSlideShow xmlSlideShow = new XMLSlideShow(is);
            is.close();
            // 获取大小
            Dimension pgsize = xmlSlideShow.getPageSize();
            // 获取幻灯片
            List<XSLFSlide> slides = xmlSlideShow.getSlides();
            StringBuffer sb = null;
            int length = slides.size();
            for (int i = 0; i < length; i++) {
                sb = new StringBuffer();
                // 解决乱码问题
                List<List<XSLFTextParagraph>> shapes = slides.get(i).getNotes().getTextParagraphs();
                for (List<XSLFTextParagraph> shape : shapes) {
                    for (XSLFTextParagraph xslfTextParagraph : shape) {
                        if (xslfTextParagraph != null && !"".equals(xslfTextParagraph.getText())) {
                            String text = xslfTextParagraph.getText();
                            if (text.indexOf("null") >= 0) {
                                text = text.replace("null", "");
                                if ("".equals(text)) {
                                    break;
                                }
                            }
                            List<XSLFTextRun> textRuns = xslfTextParagraph.getTextRuns();
                            for (XSLFTextRun xslfTextRun : textRuns) {
                                xslfTextRun.setFontFamily("宋体");
                            }
                            sb.append(xslfTextParagraph.getText());
                        }
                    }
                }
                if (sb.length() == 0) {
                    continue;
                }
                // 根据幻灯片大小生成图片
                BufferedImage img = new BufferedImage(pgsize.width * times, pgsize.height * times, BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                graphics.setPaint(Color.white);
                graphics.scale(times, times);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width * times, pgsize.height * times));
                // 最核心的代码
                slides.get(i).draw(graphics);
                // 图片将要存放的路径
                // 话术名称
                String str = sb.toString();
                str = str.replaceAll("@time=", "p").replaceAll("]", "000]");
                voiceMap.put("voice_" + (i + 1) + ".wav", str);


                String absolutePath = imgFile + "img_" + (i + 1) + ".jpeg";
                picNames.add(absolutePath);

                File jpegFile = new File(absolutePath);
                File dir = jpegFile.getParentFile();
                if (!dir.exists()) {
                    dir.mkdirs();
                }
                //如果图片存在，则不再生成
                /*if (jpegFile.exists()) {
                    continue;
                }*/
                // 这里设置图片的存放路径和图片的格式(jpeg,png,bmp等等),注意生成文件路径
                FileOutputStream out = new FileOutputStream(jpegFile);
                // 写入到图片中去
                ImageIO.write(img, "jpeg", out);
                out.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return picNames;
    }

    /**
     * 生成话术
     */
    public static void createVoice() {
        if (!voiceMap.isEmpty()) {
            for (Map.Entry<String, String> map : voiceMap.entrySet()) {
                getVoicePath(map.getKey(), pptToVideoTempFilePath, "/", map.getValue());
            }
        }
    }

    /**
     * 调用tts
     *
     * @param fileName
     * @param firstDir
     * @param secondDir
     * @param context
     */
    public static void getVoicePath(String fileName, String firstDir, String secondDir, String context) {
        try {
            String[] commandLine = new String[]{"", fileName, firstDir, secondDir, context};
            commandLine[0] = TTSFILEPATH;
            String ttsResult = commandExecutor(commandLine);
            System.out.println("tts合成结果：" + ttsResult);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * tts远程执行命令
     *
     * @param commandLine
     * @return
     */
    private static String commandExecutor(String[] commandLine) {
        String result = "";
        try {
            Runtime runtime = Runtime.getRuntime();
            Process process = runtime.exec(commandLine);
            BufferedReader stdReader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            BufferedReader stdError = new BufferedReader(new InputStreamReader(process.getErrorStream()));
            while (true) {
                String s = null;
                result += '\n';
                if ((s = stdError.readLine()) != null) {
                    result += s;
                    continue;
                }
                if ((s = stdReader.readLine()) != null) {
                    result += s;
                    continue;
                }
                break;
            }

            // 进程超时则杀掉进程，防止阻塞线程
            if (!process.waitFor(PROCESSTIMEOUT, TimeUnit.SECONDS)) {
                process.destroy();
            }
            return result;
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
            throw new RuntimeException("离线语音合成执行失败，请联系管理员！" + e.getMessage());
        }
    }


    /**
     * ssh远程连接命令执行方法
     *
     * @param command
     * @return
     */
    public static String exec(String command) {
        Runtime runtime = Runtime.getRuntime();
        System.out.println("命令为：" + command);
        String[] cmd = {"/bin/sh", "-c", command};
        try {
            Process proc = runtime.exec(cmd);
            proc.waitFor();
            // 实例化输入流，并获取网页代码
            BufferedReader reader = new BufferedReader(new InputStreamReader(proc.getErrorStream(), "UTF-8"));
            String s; // 依次循环，至到读的值为空
            StringBuilder sb = new StringBuilder();
            while ((s = reader.readLine()) != null) {
                sb.append(s);
            }
            reader.close();
            if (sb.length() > 0) {
                System.out.println("结果: " + sb.toString() + "\n");
            }
            BufferedReader stdReader = new BufferedReader(new InputStreamReader(proc.getInputStream()));
            while ((s = stdReader.readLine()) != null) {
                sb.append(s);
            }
            stdReader.close();
            if (sb.length() > 0) {
                return sb.toString();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * tts单个音频合成完整音频
     *
     * @param voiceNames
     * @return
     * @throws Exception
     */
    private static Map<String, String> combineMp3(Map<String, String> voiceNames) throws Exception {
        // 查找是否已经有合成的话术
        String cmdFind = "find " + pptToVideoTempFilePath + " -name 'voice_*.wav'";
        String existFile = exec(cmdFind);
        System.out.println("结果：" + existFile);
        if (!"".equals(existFile) && !voiceNames.isEmpty()) {
            Map<String, String> map = new LinkedHashMap<>();
            StringBuffer sb = new StringBuffer();
            StringBuffer arrSb = new StringBuffer();
            int i = 0;
            for (Map.Entry<String, String> voice : voiceNames.entrySet()) {
                // 获取文件时长
                String name = pptToVideoTempFilePath + voice.getKey();
                String cmdFfmpegTime = "ffmpeg -i " + name + " 2>&1 | grep 'Duration'| cut -d ' ' -f 4| sed s/,//";
                String ffmpegTime = exec(cmdFfmpegTime);
                // 保存话术间隔时间给图片切换用
                System.out.println("结果：" + ffmpegTime);
                map.put(voice.getKey(), ffmpegTime);
                sb.append(" -i " + pptToVideoTempFilePath + voice.getKey());
                arrSb.append("[" + i + ":0]");
                i++;
            }
            // 合并话术为一个文件
            String cmdCombine = "ffmpeg " + sb + " -filter_complex " + arrSb + "concat=n=" + (voiceNames.size()) + ":v=0:a=1[a] -map [a] " +
                    pptToVideoTempFilePath + "combined.wav";
            exec(cmdCombine);

            return map;
        }
        return null;
    }

    /**
     * 生成ppt转视频文件
     *
     * @param filePath
     * @param picNames
     * @param suffix
     * @param voiceDurations
     */
    private static void createVideo(String filePath, List<String> picNames, String suffix, Map<String, String> voiceDurations) {
        if (filePath != null && suffix != null && !picNames.isEmpty() && !voiceDurations.isEmpty()) {
            // 创建text文件，生成视频
            String cmdDeleteTmpTxt = "find " + pptToVideoTempFilePath + " -name 'temp.txt' -exec rm -f {} \\;";
            exec(cmdDeleteTmpTxt);
            String cmdCreateVideoTxt = "touch " + pptToVideoTempFilePath + "temp.txt";
            exec(cmdCreateVideoTxt);
            String tempTxtFilePath = pptToVideoTempFilePath + "temp.txt";
            StringBuffer cmdWriteTxt = new StringBuffer("");
            int i = 0;
            for (Map.Entry<String, String> mm : voiceDurations.entrySet()) {
                System.out.println(picNames.get(i) + "  ---  " + mm.getValue());
                cmdWriteTxt.append("echo 'file " + picNames.get(i) + "' >> " + tempTxtFilePath + " \n");
                cmdWriteTxt.append("echo 'duration " + mm.getValue() + "' >> " + tempTxtFilePath + " \n");
                if (i == picNames.size() - 1) {
                    cmdWriteTxt.append("echo 'file " + picNames.get(i) + "' >> " + tempTxtFilePath + " ");
                }
                i++;
            }
            // 创建生成图片的文件
            exec(cmdWriteTxt.toString());
            // 生成视频
            String cmdVideo = "ffmpeg -f concat -safe 0 -i " + pptToVideoTempFilePath + "temp.txt -vsync vfr -pix_fmt yuv420p " + pptToVideoTempFilePath + "noVoice.mp4";
            exec(cmdVideo);
            // ffmpeg提取视频
            String cmdGetAvi = "ffmpeg -i " + pptToVideoTempFilePath + "noVoice.mp4 -vcodec copy -an " + pptToVideoTempFilePath + "noVoice.avi";
            exec(cmdGetAvi);
            // 步骤二：ffmpeg替换音频命令（就是把音频视频合并起来）
            String cmdMp4 = "ffmpeg -i " + pptToVideoTempFilePath + "noVoice.avi -i " + pptToVideoTempFilePath + "combined.wav " + fileDirPath + FileName + "-pptToVideo.mp4";
            exec(cmdMp4);
        }
    }

    /**
     * 删除临时文件
     *
     * @param
     * @param
     */
    private static void deleteTempFile() {
        //操作完成删除中间临时文件
        String cmdDeleteTmpJpeg = "find " + pptToVideoTempFilePath + " -name 'img_*.jpeg' -exec rm -f {} \\;";
        exec(cmdDeleteTmpJpeg);
        String cmdDeleteTmpWav = "find " + pptToVideoTempFilePath + " -name 'voice_*.wav' -exec rm -f {} \\;";
        exec(cmdDeleteTmpWav);
        String cmdDeleteTmpCombinedWav = "find " + pptToVideoTempFilePath + " -name 'combined*.wav' -exec rm -f {} \\;";
        exec(cmdDeleteTmpCombinedWav);
        String cmdDeleteTmpMp4 = "find " + pptToVideoTempFilePath + " -name '*noVoice*' -exec rm -f {} \\;";
        exec(cmdDeleteTmpMp4);
        //删除临时目录文件
        String cmdDeleteTmpDir = "rm -rf " + pptToVideoTempFilePath + " \\";
        exec(cmdDeleteTmpDir);
    }
}
