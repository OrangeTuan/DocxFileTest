import java.util.HashMap;
import java.util.Map;

public class Main {

    public static void main(String[] args) {
        DocxExportI doc = new DocxExportImpl();
        String path = Main.class.getResource("/file/test.docx").getPath();
        String outPath = "src/file/testOut.docx";
        Map<String, String> params = new HashMap<>();
        params.put("name1","小陈");
        params.put("name2", "小红");
        params.put("sport1", "打篮球");
        params.put("sport2", "打乒乓球");
        doc.export(path, outPath, params);
    }

}
