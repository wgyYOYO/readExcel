package wgy.provider;

import com.alibaba.fastjson.JSON;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import io.swagger.annotations.ApiImplicitParam;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import wgy.action.read;
import wgy.entity.CertificateData;
import wgy.entity.CertificateField;
import wgy.entity.Response;
import wgy.entity.User;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Timestamp;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static wgy.action.DocOperator.addBlankSpace;
import static wgy.action.DocOperator.toCumstomDoc;

@RestController
@CrossOrigin
public class excelProviderImpl implements excelProvider {

//     ResponseBase test() throws Exception {
//        List<User> date = new read().getDate();
//        System.out.println(date.toString());
//    }

    @ApiOperation(value = "通过用户名查询用户信息", notes = "通过用户名查询用户信息", produces = "application/json")
//    @ApiImplicitParam(name = "name", value = "用户名", paramType = "query", required = true, dataType = "String")
    @RequestMapping(value = "/user", method = RequestMethod.GET)
    public Response<Map<String, Object>> get() {
        Response<Map<String, Object>> response = new Response<>();
        Map<String, Object> user = new HashMap<>();
        user.put("name", "demo");
        user.put("age", 25);
        response.setData(user);
        System.out.println("掉到了");
        return response;
    }

    @ApiOperation(value = "通过用户名查询用户信息", notes = "通过用户名查询用户信息", produces = "application/json")
//    @ApiImplicitParam(name = "name", value = "用户名", paramType = "query", required = true, dataType = "String")
    @RequestMapping(value = "/upload", method = RequestMethod.GET)
    public Response<String> upload() throws Exception {
        Response<String> response = new Response<>();
        List<User> date = new read().getDate();
        System.out.println("掉到了");
        response.setData(JSON.toJSONString(date));
        return response;
    }

    //    @ApiOperation(value = "通过用户名查询用户信息", notes = "通过用户名查询用户信息", produces = "application/json")
    @RequestMapping(value = "/uploadFile", method = RequestMethod.POST)
    @ResponseBody
    public Response<String> uploadFile(@RequestParam("file") MultipartFile file) throws Exception {
        Response<String> response = new Response<>();
        String fileRealName = file.getOriginalFilename();//获得原始文件名;
        int pointIndex =  fileRealName.lastIndexOf(".");//点号的位置
        String fileSuffix = fileRealName.substring(pointIndex);//截取文件后缀
        String fileNewName = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS"));//新文件名,时间戳形式yyyyMMddHHmmssSSS
        String saveFileName = fileNewName.concat(fileSuffix);//新文件完整名（含后缀）

        String filePath = "D:\\zltFile";
        List<User> date =new ArrayList<>();
        File path = new File(filePath); //判断文件路径下的文件夹是否存在，不存在则创建
        if (!path.exists()) {
            path.mkdirs();
        }
        File savedFile = new File(filePath+"\\"+saveFileName);
        boolean isCreateSuccess = savedFile.createNewFile(); // 是否创建文件成功
//        List<NiceXWPFDocument> xwpfDocuments = new ArrayList<>();
        NiceXWPFDocument newDoc = new NiceXWPFDocument();
        if (isCreateSuccess) {
            file.transferTo(savedFile);
            date = new read().getDateByFile(savedFile);
            List<User> finalDate = date;
            for (User item :finalDate) {
                XWPFTemplate template =  XWPFTemplate.compile("D:\\zltFile\\data\\template1_1.docx").render(
                        new HashMap<String, Object>(){{
                            put("姓名", item.getName());
                            put("性别", item.getSex());
                            put("学号", item.getStuId());
                        }});
                NiceXWPFDocument xwpfDocument = template.getXWPFDocument();
                newDoc.merge(xwpfDocument);
            }

//            String resultPath ="D:\\zltFile\\word\\output1_result_"+LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS"))+".docx";
//            template.writeAndClose(new FileOutputStream(resultPath));

//            xwpfDocuments.add(xwpfDocument);
        }
//        if (xwpfDocuments.size()>0){
//
//            for (int i = 0; i < xwpfDocuments.size(); i++) {
//                newDoc.merge(xwpfDocuments.get(i));
//            }
            // 生成新文档
            FileOutputStream out = new FileOutputStream("D:\\zltFile\\test\\new_doc_"+LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS"))+".docx");
            newDoc.write(out);
            newDoc.close();
            out.close();
            System.out.println("合并word成功");

//        }

//        XWPFTemplate template = XWPFTemplate.compile("D:\\zltFile\\data\\template1_1.docx").render(
//                new HashMap<String, Object>(){{
//                    put("姓名", "张三");
//                    put("性别", "男");
//                    put("学号", "171164476");
//                }});

        System.out.println(JSON.toJSONString(date));
        response.setData(JSON.toJSONString("掉到了"));
        return response;
    }
}
