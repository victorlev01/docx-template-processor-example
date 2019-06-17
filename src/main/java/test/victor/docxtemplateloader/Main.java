package test.victor.docxtemplateloader;

import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws Exception {
        Path filePath = Paths.get("C:\\Java\\test-projects\\docx-template-processor\\src\\main\\resources\\early_payment_template.docx");
        Path newFilePath = Paths.get("C:\\Java\\test-projects\\docx-template-processor\\src\\main\\resources\\generated.docx");
        Map<String, Object> formValues = new HashMap<>();
        List<Map> supplies = new ArrayList<>();
        supplies.add(new HashMap() {{
            put("name", "Первый");
            put("sum", "1000");
            put("discount", "100%");
        }});
        supplies.add(new HashMap() {{
            put("name", "Второй");
            put("sum", "6596433123");
            put("discount", "12%");
        }});
        supplies.add(new HashMap() {{
            put("name", "Последний");
            put("sum", "-1");
            put("discount", "-100%");
        }});
        formValues.put("supplies", supplies);

        formValues.put("inn", "100500");

        new Main().generateWordFile(filePath, formValues, newFilePath);
    }

    public void generateWordFile(Path filePath, Map<String, Object> formMap, Path newFilePath) throws XDocReportException, IOException {
        IXDocReport report = XDocReportRegistry.getRegistry().loadReport(Files.newInputStream(filePath), TemplateEngineKind.Freemarker);
        FieldsMetadata fieldsMetadata = report.createFieldsMetadata();
        fillMetaByList(formMap, fieldsMetadata);
        IContext context = report.createContext();
        context.putMap(formMap);
        OutputStream os = Files.newOutputStream(newFilePath);
        report.process(context, os);
    }

    private static void fillMetaByList(Map<String, Object> map, FieldsMetadata fieldsMetadata) {
        for (Map.Entry<String, Object> entry : map.entrySet()) {
            if (entry.getValue() instanceof ArrayList) {
                fieldsMetadata.addFieldAsList(entry.getKey());
            } else if (entry.getValue() instanceof HashMap) {
                fillMetaByList((Map<String, Object>) entry.getValue(), fieldsMetadata);
            }
        }
    }
}
