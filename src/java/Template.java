package java;

import javax.xml.stream.XMLStreamWriter;

public class Template {
    private String name;
    private boolean isCharacterSheet;
    private String fileName;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public boolean isCharacterSheet() {
        return isCharacterSheet;
    }

    public void setCharacterSheet(boolean characterSheet) {
        isCharacterSheet = characterSheet;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    //TODO: Check XML writing best practices.

    /**
     * Write the obkject to an XML file.
     * @param xml The XML Stream Writer
     */
    public void outputToFile(XMLStreamWriter xml) {
        
    }
}
