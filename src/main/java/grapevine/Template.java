package grapevine;

import javax.xml.bind.annotation.XmlRootElement;

@XmlRootElement
public class Template {
    private String name;
    private boolean isCharacterSheet;
    private String[] fileName;

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

    public String getFileName(String extension) {
        String lowExtension = extension.toLowerCase().trim();
        switch(lowExtension) {
            case "text":
                return fileName[0];
            case "rtf":
                return fileName[1];
            case "html":
                return fileName[2];
            default:
                return "";
        }
    }

    public void setFileName(String fileName, String extension) {
        String lowExtension = extension.toLowerCase().trim();
        switch(lowExtension) {
            case "text":
                this.fileName[0] = fileName;
            case "rtf":
                this.fileName[1] = fileName;
            case "html":
                this.fileName[2] = fileName;
        }
    }
    
}
