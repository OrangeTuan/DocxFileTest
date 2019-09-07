import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DocxExportImpl implements  DocxExportI {

    @Override
    public File export(String tempFilePath, String outFilePath, Map<String, String> params) {
        if (outFilePath == null || outFilePath.length() == 0) {
            System.out.println("生成的文件路径不能为空！");
            return null;
        }
        File tempFile = new File(tempFilePath);
        if (!tempFile.exists()) {
            System.out.println("找不到模板文件！");
            return null;
        }
        if (params == null) return null;
        try (InputStream in = new FileInputStream(tempFilePath);
             OutputStream outputStream = new FileOutputStream(outFilePath);) {
            XWPFDocument doc = new XWPFDocument(in);
            //获取所有段落
            for (Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();iterator.hasNext();){
                XWPFParagraph p = iterator.next();
                replaceInParagraph(p, params);
            }
            doc.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private void replaceInParagraph(XWPFParagraph paragraph, Map<String, String> params) {
        List<XWPFRun> runList = null;
        Matcher matcher = matcher(paragraph.getParagraphText());
        if (matcher.find()) {
            runList = paragraph.getRuns();
            for (int i = 0;i < runList.size(); i++) {
                XWPFRun run = runList.get(i);
                String runText = run.text();
                if (matcher(runText).find()){
                    while ((matcher = matcher(runText)).find()){
                        runText = matcher.replaceAll(String.valueOf(params.get(matcher.group(1))));
                    }
                    paragraph.removeRun(i);
                    paragraph.insertNewRun(i).setText(runText);
                }
            }
        }
    }

    private Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

}
