package grapevine.model;

import grapevine.constants.ExperienceChange;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.Serializable;

@XmlRootElement
public class ExperienceAward implements Serializable {
    @XmlElement
    private String name;
    @XmlElement
    private ExperienceChange change;
    @XmlElement
    private int amount;
    @XmlElement
    private String reason;
    @XmlElement
    private boolean xp;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public ExperienceChange getChange() {
        return change;
    }

    public void setChange(ExperienceChange change) {
        this.change = change;
    }

    public int getAmount() {
        return amount;
    }

    public void setAmount(int amount) {
        this.amount = amount;
    }

    public String getReason() {
        return reason;
    }

    public void setReason(String reason) {
        this.reason = reason;
    }

    public boolean isXp() {
        return xp;
    }

    public void setXp(boolean xp) {
        this.xp = xp;
    }

    /**
     * Sets text appropriate to change type
     * @return The change time concatenated with the amount changed.
     */
    public String changeTypeText() {
        String text = "unknown";
        switch(change) {
            case EARNED:
                text = "earn " + amount;
                break;
            case SPENT:
                text = "spend " + amount;
                break;
            case DEDUCTED:
                text = "deduct " + amount;
                break;
            case UNSPENT:
                text = "unspend " + amount;
                break;
            case SET_EARNED:
                text = "set earned to " + amount;
                break;
            case SET_UNSPENT:
                text = "set unspent to " + amount;
                break;
            case COMMENT:
                text = "comment";
                break;
        }
        return text;
    }

    //Output to File and From file will be handled by jaxb in a calling method. No need for methods here.

    /**
     * ToDo: Test this for backwards compatibility. A lot.
     * Writes the object to binary depending on software version
     * @param outputStream The in-progress binary output stream
     * @param version The version of the file to be written
     * @throws IOException on error
     */
    public void outputToBinary(ObjectOutputStream outputStream, double version) throws IOException {
        if (version >= 4.0) {
            outputStream.writeObject(this);
        } else {
            byte[] nameBytes = name.getBytes();
            byte[] changeBytes = change.toString().getBytes();
            byte[] reasonBytes = reason.getBytes();

            outputStream.writeBoolean(xp);
            outputStream.writeInt(nameBytes.length);
            outputStream.write(nameBytes);
            outputStream.writeInt(changeBytes.length);
            outputStream.write(changeBytes);
            outputStream.writeInt(amount);
            outputStream.writeInt(reasonBytes.length);
            outputStream.write(reasonBytes);
        }
    }

    /**
     * ToDo: Test this for backwards compatibility. A lot.
     * Reads the object from binary depending on software version
     * @param inputStream The in-progress binary input stream
     * @param version The version of the file to be written
     * @throws IOException on error
     * @throws ClassNotFoundException on error
     */
    public static ExperienceAward inputFromBinary(ObjectInputStream inputStream, double version) throws IOException, ClassNotFoundException {
        ExperienceAward ea = new ExperienceAward();
        int length;
        if (version >= 4.0) {
            ea =  (ExperienceAward) inputStream.readObject();
        } else {
            ea.setXp(inputStream.readBoolean());
            length = inputStream.readInt();
            ea.setName(new String(inputStream.readNBytes(length)));
            length = inputStream.readInt();
            ea.setChange(ExperienceChange.valueOf(new String(inputStream.readNBytes(length))));
            ea.setAmount(inputStream.readInt());
            length = inputStream.readInt();
            ea.setReason(new String(inputStream.readNBytes(length)));
        }
        return ea;
    }
}
