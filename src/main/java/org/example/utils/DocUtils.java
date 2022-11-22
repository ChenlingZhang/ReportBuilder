package org.example.utils;

import com.aspose.words.DocumentBuilder;
import lombok.extern.slf4j.Slf4j;

import java.io.File;


@Slf4j
public class DocUtils {
    /**
     * 用于生成文章1-4级的标题
     * @param builder DocumentBuilder
     * @param title 需要生成的文章标题
     * @param titleLevel 需要生成的标题字号1-4
     * @throws Exception 统一错误抛出
     */
    public static void generateTitle(DocumentBuilder builder, String title, int titleLevel) throws Exception {
        switch (titleLevel){
            case 1:
                builder.insertHtml("<h1 style='text-align:left;font-family:Simsun;'>" + title + "</h1>");
                log.info("生成1级标题{}", title);
                break;
            case 2:
                builder.insertHtml("<h2 style='text-align:left;font-family:Simsun;'>" + title + "</h2>");
                log.info("生成2级标题{}", title);
                break;
            case 3:
                builder.insertHtml("<h3 style='text-align:left;font-family:Simsun;'>" + title + "</h3>");
                log.info("生成3级标题{}", title);
                break;
            case 4:
                builder.insertHtml("<h4 style='text-align:left;font-family:Simsun;'>" + title + "</h4>");
                log.info("生成4级标题{}", title);
                break;

        }
    }

    /**
     * 生成文件
     * @param path 传入路径+文件名例如 /doc/test.docx
     * @return
     * @throws Exception
     */
    public static File createFile(String path) throws Exception {
        File file = new File(path);
        if (file.exists() || !file.isFile()) {
            throw new Exception("文件创建失败，路径中已存在改文件或传入的不是一个文件");
        }
        file.createNewFile();
        return file;
    }


}
