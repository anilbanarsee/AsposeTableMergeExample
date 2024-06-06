package service;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.ReplaceAction;

public class Main {

  public static void main(String[] args) {

    mergeTable("template.docx");
    mergeTable("template.odt");
    mergeTable("template.rtf");
  }

  // Replaces [[FIELD_1]] with a table
  private static void mergeTable(String template) {
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

      doc.save("out-docs/%s".formatted(template));
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }
}
