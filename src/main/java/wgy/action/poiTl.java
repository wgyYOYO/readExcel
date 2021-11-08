package wgy.action;

import com.deepoove.poi.XWPFTemplate;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

public class poiTl {
    @Test
    public void test() throws IOException {
        XWPFTemplate template = XWPFTemplate.compile("D:\\zltFile\\data\\template1_1.docx").render(
                new HashMap<String, Object>(){{
                    put("姓名", "张三");
                    put("性别", "男");
                    put("学号", "171164476");
                }});
        template.writeAndClose(new FileOutputStream("D:\\zltFile\\word\\output1_result.docx"));
    }
}
