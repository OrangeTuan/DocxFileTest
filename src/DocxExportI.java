import java.io.File;
import java.util.Map;

public interface DocxExportI {
    public File export(String tempFilePath, String outFilePath, Map<String, String> params);
}
