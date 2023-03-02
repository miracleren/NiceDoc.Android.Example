package com.miracleren.testdoc;

import androidx.annotation.NonNull;
import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

import android.Manifest;
import android.app.Activity;

import android.content.pm.PackageManager;

import android.os.Build;
import android.os.Bundle;
import android.widget.TextView;

import com.miracleren.NiceDoc;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;


public class MainActivity extends AppCompatActivity {


    @RequiresApi(api = Build.VERSION_CODES.R)
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        //引用poishadow必须添加
        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl");

        getPermissionCamera(this);
        String docTempPath = getExternalCacheDir().getPath() + "/";
        try {
            InputStream docTem = this.getAssets().open("test.docx");
            writeToFile(docTempPath + "test.docx", docTem);
            InputStream pngTem = this.getAssets().open("head.png");
            writeToFile(docTempPath + "head.png", pngTem);


            //测试示例模板生成
            NiceDoc docx = new NiceDoc(docTempPath + "test.docx");

            Map<String, Object> labels = new HashMap<>();
            //值标签
            labels.put("startTime", "1881年9月25日");
            labels.put("endTime", "1936年10月19日");
            labels.put("title", "精选作品目录");
            labels.put("press", "鲁迅同学出版社");

            //枚举标签
            labels.put("likeBook", 2);
            //布尔标签
            labels.put("isQ", true);
            //等于
            labels.put("isNew", 2);
            //多选二进制值
            labels.put("look", 3);
            //if语句
            labels.put("showContent", 2);
            //日期格式标签
            labels.put("printDate", new Date());

            labels.put("fileReceiveBy", "陈先生");
            labels.put("fileRelation", 2);
            labels.put("fileDate", new Date());

            //添加头像
            labels.put("headImg", docTempPath + "head.png");

            docx.pushLabels(labels);

            //表格
            List<Map<String, Object>> books = new ArrayList<>();
            Map<String, Object> book1 = new HashMap<>();
            book1.put("name", "汉文学史纲要");
            book1.put("time", "1938年，鲁迅全集出版社");
            books.add(book1);
            Map<String, Object> book2 = new HashMap<>();
            book2.put("name", "中国小说史略");
            book2.put("time", "1923年12月，上册；1924年6月，下册");
            books.add(book2);
            docx.pushTable("books", books);


            //生成文档
            String docName = UUID.randomUUID() + ".docx";
            docx.save(docTempPath, docName);

            TextView view = findViewById(R.id.testInfo);
            view.setText("原模板文档：" + docTempPath + "test.docx" + "\n\n成功生成文档：" + docTempPath + docName);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void getPermissionCamera(Activity activity) {
        int readPermissionCheck = ContextCompat.checkSelfPermission(activity, Manifest.permission.READ_EXTERNAL_STORAGE);
        int writePermissionCheck = ContextCompat.checkSelfPermission(activity, Manifest.permission.WRITE_EXTERNAL_STORAGE);
        if (readPermissionCheck != PackageManager.PERMISSION_GRANTED || writePermissionCheck != PackageManager.PERMISSION_GRANTED) {
            String[] permissionList = new String[]{Manifest.permission.READ_EXTERNAL_STORAGE,
                    Manifest.permission.WRITE_EXTERNAL_STORAGE,
                    Manifest.permission.ACCESS_NETWORK_STATE};
            ActivityCompat.requestPermissions(
                    activity,
                    permissionList,
                    0);
        }
    }

    private static void writeToFile(String path, InputStream input)
            throws IOException {
        int index;
        byte[] bytes = new byte[1024];
        FileOutputStream newFile = new FileOutputStream(path);
        while ((index = input.read(bytes)) != -1) {
            newFile.write(bytes, 0, index);
            newFile.flush();
        }
        newFile.close();
        input.close();
    }

}