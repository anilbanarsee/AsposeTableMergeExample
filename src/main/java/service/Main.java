package service;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.ReplaceAction;

import java.io.File;
import java.util.Arrays;
import java.util.List;

public class Main {

    private static final List<String> FORMATS = List.of("docx", "odt", "rtf");

    public static void main(String[] args) {

        // mergeAllFilesInDir("src/main/resources")

        mergeIntoAllFormats("watermark.docx");
    }

    private static void mergeIntoAllFormats(String template) {
        FORMATS.stream()
                .map(ext -> changeExtension(template, ext))
                .forEach(output -> mergeTable(template, output));
    }

    private static void mergeAllFilesInDir(String directory) {
        var dir = new File(directory);
        var files = dir.listFiles();
        Arrays.stream(files != null ? files : new File[0])
                .filter(File::isFile)
                .map(File::getName)
                .forEach(Main::mergeTable);
    }

    private static void mergeTable(String template) {
        mergeTable(template, template);
    }

    // Replaces [[FIELD_1]] with a table
    private static void mergeTable(String template, String output) {
        try {
            var doc = new Document("src/main/resources/%s".formatted(template));

            FindReplaceOptions options = new FindReplaceOptions();
            options.setReplacingCallback(replacingArgs -> {
                var builder = new DocumentBuilder(doc);
                builder.moveTo(replacingArgs.getMatchNode());
                builder.startTable();

                builder.insertCell();
                builder.write("Hickory Dickory Dock");
                builder.insertCell();
                builder.write("The");
                builder.insertCell();
                builder.write("Mouse");
                builder.endRow();

                builder.insertCell();
                builder.write("Ran Up");
                builder.insertCell();
                builder.write("The");
                builder.insertCell();
                builder.write("Clock");
                builder.endRow();

                replacingArgs.setReplacement("");
                return ReplaceAction.REPLACE;
            });

            doc.getRange().replace("[[FIELD_1]]", "", options);

            doc.save("out-docs/%s".formatted(output));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static String changeExtension(String filename, String extension) {
        return "%s.%s".formatted(filename.split("\\.")[0], extension);
    }
}
